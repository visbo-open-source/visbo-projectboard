Imports System.Math

Public Class clsProjekt
    Inherits clsProjektvorlage

    ' diese Variable würde die Variable aus der inherited Klasse clsProjektvorlage überschatten .. 
    ' deshalb auskommentiert 
    'Private _Dauer As Integer


    'Private AllPhases As List(Of clsPhase)
    Private relStart As Integer
    Private imarge As Double
    Private uuid As Long
    Private iDauer As Integer
    Private _StartOffset As Integer
    Private _Start As Integer
    Private _earliestStart As Integer
    Private _latestStart As Integer
    Private _Status As String
    Private _earliestStartDate As Date
    Private _startDate As Date
    Private _latestStartDate As Date
    Private _ampelStatus As Integer
    Private _ampelErlaeuterung As String
    Private _name As String

    ' geändert 07.04.2014: Damit jedes Projekt auf der Projekttafel angezeigt werden kann.
    Private NullDatum As Date = StartofCalendar



    ' Deklarationen der Events 
    Public Property shpUID As String

    Public Property Risiko As Double
    Public Property StrategicFit As Double
    Public Property Erloes As Double
    Public Property leadPerson As String
    'Public Property tfSpalte As Integer
    Public Property tfZeile As Integer
    Public Property variantName As String
    Public Property Id As String
    Public Property timeStamp As Date

    ' ergänzt am 26.10.13 - nicht in Vorlage aufgenommen, da es für jedes Projekt individuell ist 
    Public Property description As String
    Public Property volume As Double
    Public Property complexity As Double
    Public Property businessUnit As String

    ' ergänzt am 30.1.14 - diffToPrev , wird benutzt, um zu kennzeichnen , welches Projekt sich im Vergleich zu vorher verändert hat 
    Public Property diffToPrev As Boolean

    ''' <summary>
    ''' synchronisiert die Arrays mit der evtl veränderten Array Länge durch eine Verschiebung des Projekts 
    ''' berechnet und bestimmt die XWerte der Rollen und Kostenarten für die Phasen neu
    ''' wird aus set Startdate heraus aufgerufen; dadurch kann es sein, daß sich die 
    ''' Dimension der Arrays für die Rollen und kostenarten verändert 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub syncXWertePhases()
        Dim tmpValue As Boolean = True
        Dim cphase As clsPhase
        Dim dimension As Integer
        Dim phaseStart As Date, phaseEnd As Date
        Dim notYetDone As Boolean = True


        ' prüfen, ob die Gesamtlänge übereinstimmt  
        For p = 1 To Me.CountPhases
            cphase = Me.getPhase(p)
            phaseEnd = cphase.getEndDate
            phaseStart = cphase.getStartDate

            dimension = getColumnOfDate(phaseEnd) - getColumnOfDate(phaseStart)

            If cphase.CountRoles > 0 Then

                ' hier müssen jetzt die Xwerte neu gesetzt werden 
                Call cphase.calcNewXwerte(dimension, 1)
                notYetDone = False

            End If

            If cphase.CountCosts > 0 And notYetDone Then

                ' hier müssen jetzt die Xwerte neu gesetzt werden 
                Call cphase.calcNewXwerte(dimension, 1)

            End If


        Next




    End Sub

    ''' <summary>
    ''' setzt den Namen des Projektes fest oder gibt ihn zurück
    ''' gleichzeitig wird auch der Name der Phase(1), sofern sie bereits existiert, auf diesen Namen festgesetzt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property name As String
        Get
            name = _name
        End Get

        Set(value As String)

            Try
                If value.Trim.Length > 0 Then
                    _name = value.Trim
                    
                Else
                    _name = ""
                End If

            Catch ex As Exception
                _name = ""
            End Try

            If Me.CountPhases > 0 Then
                Me.getPhase(1).name = _name
            End If


        End Set
    End Property


    Public ReadOnly Property isConsistent As Boolean

        Get
            Dim tmpValue As Boolean = True
            Dim cphase As clsPhase
            Dim dimension As Integer
            Dim phaseStart As Date, phaseEnd As Date


            If Me.Dauer <> getColumnOfDate(Me.endeDate) - getColumnOfDate(Me.startDate) + 1 Then
                tmpValue = False
            End If

            ' prüfen, ob die Gesamtlänge übereinstimmt  
            For p = 1 To Me.CountPhases
                cphase = Me.getPhase(p)
                phaseEnd = cphase.getEndDate
                phaseStart = cphase.getStartDate

                dimension = getColumnOfDate(phaseEnd) - getColumnOfDate(phaseStart)

                For r = 1 To cphase.CountRoles
                    If dimension <> cphase.getRole(r).Xwerte.Length - 1 Then
                        tmpValue = False
                    End If
                Next

                For k = 1 To cphase.CountCosts
                    If dimension <> cphase.getCost(k).Xwerte.Length - 1 Then
                        tmpValue = False
                    End If
                Next

            Next

            isConsistent = tmpValue

        End Get

    End Property

    Public Overrides Sub AddPhase(ByVal phase As clsPhase)

        Dim phaseEnde As Double
        Dim maxM As Integer

        With phase

            phaseEnde = .startOffsetinDays + .dauerInDays - 1

            For m = 1 To .CountResults
                If phaseEnde < .startOffsetinDays + .getResult(m).offset Then
                    phaseEnde = .startOffsetinDays + .getResult(m).offset
                End If
            Next

        End With

        If phaseEnde > 0 Then

            maxM = DateDiff(DateInterval.Month, Me.startDate, Me.startDate.AddDays(phaseEnde)) + 1
            If maxM <> _Dauer And maxM > 0 Then
                _Dauer = maxM
                ' hier muss jetzt die Dauer der Allgemeinen Phase angepasst werden ... 
            End If
        End If


        AllPhases.Add(phase)


    End Sub

    ''' <summary>
    ''' Methode prüft auf Identität mit einem Vergleichsprojekt 
    ''' type 0 (Overview) prüft auf: 
    ''' Startdatum, Phasen, Milestones, Personalkosten, Sonstige Kosten, Ergebnis, Attribute, Projekt-Ampel, Milestone-Ampeln
    ''' type 1 (strong role identity) prüft, welche Rollen unterschiedliche Bedarfe in den Monaten haben
    ''' type 2 (weak role identity) prüft, ob die Gesamt-Summen jeweils identisch / unterschiedlich sind
    ''' type 3 (strong cost identity) prüft, in welchen Kostenarten unterschiedliche Bedarfe in den Monaten sind
    ''' type 4 (weak cost identity) prüft, ob die Gesamt-Summen jeweils identisch / unterschiedlich sind
    ''' </summary>
    ''' <param name="vglproj">Projekt vom Typ clsProjekt</param>
    ''' <param name="absolut">soll absolut verglichen werden oder relativ; nur relevant bei Overview</param>
    ''' <param name="type">gibt den Vergleichstyp an</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property listOfDifferences(ByVal vglproj As clsProjekt, ByVal absolut As Boolean, ByVal type As Integer) As Collection
        Get
            Dim isDifferent As Boolean = False
            Dim tmpCollection As New Collection
            Dim hValues() As Double, cValues() As Double
            Dim hdates As SortedList(Of Date, String)
            Dim cdates As SortedList(Of Date, String)

            Dim verify As Integer = Me.dauerInDays
            verify = vglproj.dauerInDays


            Select Case type

                Case 0 ' Overview

                    ' Ist das startdatum unterschiedlich?
                    If Me.startDate.Date <> vglproj.startDate.Date Then
                        Try
                            tmpCollection.Add(PThcc.startdatum, PThcc.startdatum)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Phasen identisch sind 
                    hValues = Me.getPhaseInfos
                    cValues = vglproj.getPhaseInfos
                    If arraysAreDifferent(hValues, cValues) Then
                        Try
                            tmpCollection.Add(PThcc.phasen, PThcc.phasen)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Milestones identisch sind 
                    hdates = Me.getMilestones
                    cdates = vglproj.getMilestones
                    If dateListsareDifferent(hdates, cdates) Then
                        Try
                            tmpCollection.Add(PThcc.resultdates, PThcc.resultdates)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen , ob die Personalkosten identisch sind 
                    hValues = Me.getAllPersonalKosten
                    cValues = vglproj.getAllPersonalKosten
                    If arraysAreDifferent(hValues, cValues) And (hValues.Sum > 0 Or cValues.Sum > 0) Then
                        Try
                            tmpCollection.Add(PThcc.perscost, PThcc.perscost)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob sonstige Kosten identisch sind 
                    hValues = Me.getGesamtAndereKosten
                    cValues = vglproj.getGesamtAndereKosten
                    If arraysAreDifferent(hValues, cValues) And (hValues.Sum > 0 Or cValues.Sum > 0) Then
                        Try
                            tmpCollection.Add(PThcc.othercost, PThcc.othercost)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob das Ergebnis identisch ist 
                    Dim aktBudget As Double, aktPCost As Double, aktSCost As Double, aktRCost As Double, aktErg As Double
                    Dim vglBudget As Double, vglPCost As Double, vglSCost As Double, vglRCost As Double, vglErg As Double

                    With Me
                        .calculateRoundedKPI(aktBudget, aktPCost, aktSCost, aktRCost, aktErg)
                    End With

                    With vglproj
                        .calculateRoundedKPI(vglBudget, vglPCost, vglSCost, vglRCost, vglErg)
                    End With

                    If aktErg <> vglErg Then
                        Try
                            tmpCollection.Add(PThcc.ergebnis, PThcc.ergebnis)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Attribute identisch sind
                    If Me.StrategicFit <> vglproj.StrategicFit Or _
                        Me.Risiko <> vglproj.Risiko Then
                        Try
                            tmpCollection.Add(PThcc.fitrisk, PThcc.fitrisk)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Projekt Ampel unterschiedlich ist 
                    If Me.ampelStatus <> vglproj.ampelStatus Then
                        Try
                            tmpCollection.Add(PThcc.projektampel, PThcc.projektampel)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Meilenstein Ampeln unterschiedlich sind 
                    hValues = Me.getMilestoneColors
                    cValues = vglproj.getMilestoneColors
                    If arraysAreDifferent(hValues, cValues) Then
                        Try
                            tmpCollection.Add(PThcc.resultampel, PThcc.resultampel)
                        Catch ex As Exception

                        End Try

                    End If

                Case 1 ' strong role identity
                    Dim hUsedRoles As Collection = Me.getUsedRollen
                    Dim cUsedRoles As Collection = vglproj.getUsedRollen

                    For Each role In hUsedRoles

                        
                        hValues = Me.getRessourcenBedarf(role)
                        If cUsedRoles.Contains(role) Then

                            cValues = vglproj.getRessourcenBedarf(role)
                            If arraysAreDifferent(hValues, cValues) And (hValues.Sum > 0 Or cValues.Sum > 0) Then
                                Try
                                    tmpCollection.Add(role, role)
                                Catch ex As Exception

                                End Try
                            End If
                        Else
                            If hValues.Sum > 0 Then
                                Try
                                    tmpCollection.Add(role, role)
                                Catch ex As Exception

                                End Try
                            End If

                        End If
                        
                    Next

                    ' jetzt muss noch geprüft werden, ob es in vglproj Rollen gibt, die nicht in hproj enthalten sind 
                    ' die müssen dann auf alle fälle aufgenommen werden 

                    For Each role In cUsedRoles

                        cValues = vglproj.getRessourcenBedarf(role)

                        If Not hUsedRoles.Contains(role) And cValues.Sum > 0 Then
                            Try
                                tmpCollection.Add(role, role)
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                Case 2 ' weak role identity
                    Dim hUsedRoles As Collection = Me.getUsedRollen
                    Dim cUsedRoles As Collection = vglproj.getUsedRollen
                    ReDim hValues(0)
                    ReDim cValues(0)

                    For Each role In hUsedRoles
                        hValues(0) = Me.getRessourcenBedarf(role).Sum

                        If cUsedRoles.Contains(role) Then

                            cValues(0) = vglproj.getRessourcenBedarf(role).Sum
                            If hValues(0) <> cValues(0) Then
                                Try
                                    tmpCollection.Add(role, role)
                                Catch ex As Exception

                                End Try
                            End If
                        ElseIf hValues(0) > 0 Then
                            Try
                                tmpCollection.Add(role, role)
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                    ' jetzt muss noch geprüft werden, ob es in vglproj Rollen gibt, die nicht in hproj enthalten sind 
                    ' die müssen dann auf alle fälle aufgenommen werden 

                    For Each role In cUsedRoles

                        cValues(0) = vglproj.getRessourcenBedarf(role).Sum

                        If Not hUsedRoles.Contains(role) And cValues(0) > 0 Then
                            Try
                                tmpCollection.Add(role, role)
                            Catch ex As Exception

                            End Try

                        End If

                    Next

                Case 3 ' strong cost identity

                    Dim hUsedCosts As Collection = Me.getUsedKosten
                    Dim cUsedCosts As Collection = vglproj.getUsedKosten

                    For Each cost In hUsedCosts
                        hValues = Me.getKostenBedarf(cost)

                        If cUsedCosts.Contains(cost) Then

                            cValues = vglproj.getKostenBedarf(cost)
                            If arraysAreDifferent(hValues, cValues) And (hValues.Sum > 0 Or cValues.Sum > 0) Then
                                Try
                                    tmpCollection.Add(cost, cost)
                                Catch ex As Exception

                                End Try
                            End If
                        ElseIf hValues.Sum > 0 Then
                            Try
                                tmpCollection.Add(cost, cost)
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                    ' jetzt muss noch geprüft werden, ob es in vglproj Rollen gibt, die nicht in hproj enthalten sind 
                    ' die müssen dann auf alle fälle aufgenommen werden 

                    For Each cost In cUsedCosts
                        cValues = vglproj.getKostenBedarf(cost)
                        If Not hUsedCosts.Contains(cost) And cValues.Sum > 0 Then
                            Try
                                tmpCollection.Add(cost, cost)
                            Catch ex As Exception

                            End Try
                        End If

                    Next


                Case 4 ' weak cost identity
                    Dim hUsedCosts As Collection = Me.getUsedKosten
                    Dim cUsedCosts As Collection = vglproj.getUsedKosten
                    ReDim hValues(0)
                    ReDim cValues(0)

                    For Each cost In hUsedCosts
                        hValues(0) = Me.getKostenBedarf(cost).Sum
                        If cUsedCosts.Contains(cost) Then

                            cValues(0) = vglproj.getKostenBedarf(cost).Sum
                            If arraysAreDifferent(hValues, cValues) And (hValues(0) > 0 Or cValues(0) > 0) Then
                                Try
                                    tmpCollection.Add(cost, cost)
                                Catch ex As Exception

                                End Try
                            End If
                        ElseIf hValues(0) > 0 Then
                            Try
                                tmpCollection.Add(cost, cost)
                            Catch ex As Exception

                            End Try
                        End If

                    Next

                    ' jetzt muss noch geprüft werden, ob es in vglproj Rollen gibt, die nicht in hproj enthalten sind 
                    ' die müssen dann auf alle fälle aufgenommen werden 

                    For Each cost In cUsedCosts
                        cValues(0) = vglproj.getKostenBedarf(cost).Sum
                        If Not hUsedCosts.Contains(cost) And cValues(0) > 0 Then
                            Try
                                tmpCollection.Add(cost, cost)
                            Catch ex As Exception

                            End Try

                        End If

                    Next

            End Select

            listOfDifferences = tmpCollection
        End Get
    End Property

    

    ''' <summary>
    ''' liefert zu einem gegebenen Meilenstein das definierte Datum zurück
    ''' die Ampelfarbe wird ebenfalls in das Datum als Ablauf von Sekunden nach Mitternacht integriert
    ''' 0-nicht bewertet, 1-grün, 2-gelb, 3-rot
    ''' Fehler, wenn Meilenstein nicht existiert
    ''' Existieren mehrere Meilensteine desselben Namens so wird nur der erste zurückgebracht 
    ''' </summary>
    ''' <param name="milestoneName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneDate(ByVal milestoneName As String) As Date
        Get
            Dim found As Boolean = False
            Dim cphase As clsPhase
            Dim cresult As clsResult
            Dim tmpDate As Date
            Dim p As Integer = 1
            Dim colorIndex As Integer


            Do While p <= Me.CountPhases And Not found

                cphase = Me.getPhase(p)

                cresult = cphase.getResult(milestoneName)

                If Not IsNothing(cresult) Then
                    
                    colorIndex = cresult.getBewertung(1).colorIndex
                    tmpDate = cresult.getDate.Date

                    ' jetzt wird die Ampelfarbe ins Datum kodiert 
                    tmpDate = tmpDate.AddSeconds(colorIndex)
                    found = True

                    ' jetzt wird in das Datum kodiert, ob der Meilenstein abgeschlossen sein sollte
                    ' wenn timestamp nach dem Meilenstein-Datum steht, sollte der Meilenstein abgeschlossen sein 
                    If DateDiff(DateInterval.Day, Me.timeStamp, tmpDate) < 0 Then

                        ' Meilenstein Datum liegt vor dem Datum, an dem dieser Planungs-Stand abgegeben wurde
                        tmpDate = tmpDate.AddHours(6)

                    End If

                End If

                p = p + 1

            Loop

            If found Then
                getMilestoneDate = tmpDate
            Else
                Throw New Exception("Meilenstein existiert nicht")
            End If

        End Get
    End Property

    ''' <summary>
    ''' diese Routine berücksichtigt, wieviel von der phase im Start- bzw End Monat liegt; 
    ''' es wird für Start und Ende Monat nicht automatisch 1 gesetzt, sondern ein anteiliger Wert, der sich daran bemisst, 
    ''' wieviel Phase im Start bzw End Monat liegt; 
    '''   
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>da eine Vorlage kein StartDatum kennt, muss diese Methode als overridable/overrides definiert werden   
    ''' </remarks>
    Public Overrides ReadOnly Property getPhasenBedarf(phaseName As String) As Double()

        Get
            Dim phaseValues() As Double
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer
            Dim phase As clsPhase
            Dim phaseStart As Date, phaseEnd As Date
            Dim numberOfDays As Integer
            Dim anteil As Double


            ReDim phaseValues(_Dauer - 1)

            If _Dauer > 0 Then



                anzPhasen = AllPhases.Count
                If anzPhasen > 0 Then

                    For p = 0 To anzPhasen - 1
                        phase = AllPhases.Item(p)


                        If phase.name = phaseName Then


                            phaseStart = phase.getStartDate
                            phaseEnd = phase.getEndDate


                            With phase
                                For i = 0 To .relEnde - .relStart

                                    If i = 0 Then

                                        numberOfDays = Max(0.0, DateDiff(DateInterval.Day, phaseStart, StartofCalendar.AddMonths(Me.Start + .relStart - 1).AddDays(-1)))
                                        anteil = numberOfDays / 365 * 12
                                        phaseValues(.relStart - 1 + i) = Min(1.0, anteil)

                                    ElseIf i = .relEnde - .relStart Then

                                        numberOfDays = Max(0.0, DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start + .relEnde - 2), phaseEnd))
                                        anteil = numberOfDays / 365 * 12
                                        phaseValues(.relStart - 1 + i) = Min(1.0, anteil)

                                    Else

                                        phaseValues(.relStart - 1 + i) = 1

                                    End If

                                Next
                            End With

                        End If

                    Next p ' Loop über alle Phasen
                Else
                    Throw New ArgumentException("Projekt hat keine Phasen")
                End If


                getPhasenBedarf = phaseValues

            Else
                Throw New ArgumentException("Projekt hat keine Dauer")
                getPhasenBedarf = phaseValues
            End If
        End Get

    End Property


    Public Overrides ReadOnly Property dauerInDays As Integer

        Get
            Dim i As Integer
            Dim max As Double = 0
            Dim offsetProjStart As Integer = DateDiff(DateInterval.Day, StartofCalendar, Me.startDate)

            ' Bestimmung der Dauer 

            For i = 1 To Me.CountPhases

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays Then
                        max = .startOffsetinDays + .dauerInDays
                    End If

                    ' Änderung 16.1.2014 es wird in phase.add(result) sichergestellt, daß kein Meilenstein nach Projektende, vor Projekt-Start sein kann 
                    'For m = 1 To .CountResults
                    '    If max < .startOffsetinDays + .getResult(m).offset Then
                    '        max = .startOffsetinDays + .getResult(m).offset
                    '    End If
                    'Next

                End With

            Next i

            ' jetzt aus Konsistenzgründen die Dauer in Monaten setzen 
            _Dauer = getColumnOfDate(StartofCalendar.AddDays(offsetProjStart + max - 1)) - getColumnOfDate(StartofCalendar.AddDays(offsetProjStart)) + 1

            dauerInDays = CInt(max)
            

        End Get
    End Property


    'Public Overrides ReadOnly Property Dauer() As Integer


    '    Get
    '        Dim i As Integer
    '        Dim max As Double = 0
    '        Dim maxM As Integer

    '        ' neue Bestimmung der Dauer 

    '        For i = 1 To AllPhases.Count

    '            With Me.getPhase(i)

    '                If max < .startOffsetinDays + .dauerInDays - 1 Then
    '                    max = .startOffsetinDays + .dauerInDays - 1
    '                End If

    '                ' Änderung 16.1.2014: Meilensteine wirken nicht Dauer-Verlängernd ! 
    '                ' ausserdem wird in phase.add(result) sichergestellt , dass kein Meilenstein vor Projektstart 
    '                ' bzw. nach Projektende ist 
    '                'For m = 1 To .CountResults
    '                '    If max < .startOffsetinDays + .getResult(m).offset Then
    '                '        max = .startOffsetinDays + .getResult(m).offset
    '                '    End If
    '                'Next

    '            End With

    '        Next i

    '        maxM = DateDiff(DateInterval.Month, startDate, startDate.AddDays(max)) + 1


    '        If maxM <> _Dauer Then
    '            _Dauer = maxM
    '        End If


    '        Dauer = _Dauer


    '    End Get

    'End Property


    Public ReadOnly Property tfspalte As Integer
        Get
            tfspalte = _Start
        End Get
    End Property

    Public Property ampelStatus As Integer
        Get
            ampelStatus = _ampelStatus
        End Get

        Set(value As Integer)
            If Not (IsNothing(value)) Then
                If IsNumeric(value) Then
                    If value >= 0 And value <= 3 Then
                        _ampelStatus = value
                    Else
                        Throw New ArgumentException("unzulässiger Ampel-Wert")
                    End If
                Else
                    Throw New ArgumentException("nicht-numerischer Ampel-Wert")
                End If
            Else
                ' ohne Bewertung
                _ampelStatus = 0
            End If

        End Set
    End Property

    Public Property ampelErlaeuterung As String
        Get
            ampelErlaeuterung = _ampelErlaeuterung
        End Get
        Set(value As String)
            If Not (IsNothing(value)) Then
                _ampelErlaeuterung = CStr(value)
            Else
                _ampelErlaeuterung = " "
            End If
        End Set
    End Property

    ''' <summary>
    ''' liefert das Ende-Datum des Projektes zurück - Readonly 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property endeDate As Date
        Get
            endeDate = Me.startDate.AddDays(Me.dauerInDays - 1).Date
        End Get
    End Property


    Public Property startDate As Date
        Get
            startDate = _startDate
        End Get

        Set(newValue As Date)
            Dim value As Date = newValue.Date
            Dim olddate As Date = _startDate
            Dim differenzInTagen As Integer = DateDiff(DateInterval.Day, olddate, value)
            Dim updatePhases As Boolean = False

            ' Änderung am 25.5.14: es ist nicht mehr erlaubt, das Startdatum innerhalb des gleichen Monats zu verschieben 
            ' es muss geprüft werden, ob es noch im Planungs-Stadium ist: nur dann darf noch verschoben werden ...
            If _Status = ProjektStatus(0) And differenzInTagen <> 0 Then
                _startDate = value
                _Start = DateDiff(DateInterval.Month, StartofCalendar, value) + 1
                ' Änderung 25.5 die Xwerte müssen jetzt synchronisert werden 
                currentConstellation = ""
               
            ElseIf _startDate = NullDatum Then
                _startDate = value
                _Start = DateDiff(DateInterval.Month, StartofCalendar, value) + 1
                If differenzInTagen <> 0 Then
                    ' mit diesem Vorgang wird die Konstellation (= Projekt-Portfolio) geändert , deshalb muss das zurückgesetzt werden 
                    currentConstellation = ""
                End If
            ElseIf _Status <> ProjektStatus(0) Then
                Throw New ArgumentException("der Startzeitpunkt kann nicht mehr verändert werden ... ")

            ElseIf DateDiff(DateInterval.Day, StartofCalendar, newValue) < 0 Then
                Throw New ArgumentException("der Startzeitpunkt liegt vor dem Kalenderstart  ... ")

            End If

            ' um _Dauer neu zu berechnen ; ergänzt am 12.5.2014
            If differenzInTagen <> 0 Then
                Dim anzahlTage As Integer = Me.dauerInDays
            End If

        End Set
    End Property


    Public Property earliestStartDate As Date
        Get
            earliestStartDate = _earliestStartDate
        End Get
        Set(value As Date)

            _earliestStartDate = value.Date


        End Set
    End Property
    ''' <summary>
    ''' wird benutzt beim Einlesen vom File, Konsistenz mit Status vorausgesetzt
    ''' </summary>
    ''' <param name="anyway"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property earliestStartDate(anyway As Boolean) As Date
        Get
            earliestStartDate = _earliestStartDate
        End Get
        Set(value As Date)
            Dim Heute As Date = Now

            _earliestStartDate = value.Date


        End Set
    End Property

    Public Property latestStartDate As Date
        Get
            latestStartDate = _latestStartDate
        End Get
        Set(value As Date)
            Dim heute As Date = Now

            _latestStartDate = value.Date

        End Set
    End Property

    ''' <summary>
    ''' wird eingesetzt beim Einlesen vom File - Konsistenz vorausgesetzt 
    ''' </summary>
    ''' <param name="anyway"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property latestStartDate(anyway As Boolean) As Date
        Get
            latestStartDate = _latestStartDate
        End Get
        Set(value As Date)
            Dim heute As Date = Now

            _latestStartDate = value.Date


        End Set
    End Property

    Public ReadOnly Property withinTimeFrame(selectionType As Integer, von As Integer, bis As Integer) As Collection
        Get
            Dim tmpListe As New Collection
            ' selection type wird aktuell noch ignoriert .... 
            Dim cphase As clsPhase


            For i = 1 To AllPhases.Count

                cphase = Me.getPhase(i)

                If Me._Start + cphase.relStart - 1 > bis Or _
                    Me._Start + cphase.relEnde - 1 < von Then
                    ' nichts tun 
                Else
                    ' ist innerhalb des Zeitrahmens
                    Try
                        tmpListe.Add(cphase.name, cphase.name)
                    Catch ex As Exception
                        ' in diesem Fall muss keine Fehlerbehandlung geamcht werden 
                        ' jede Phase wird nur einmal eingetragen ....

                    End Try

                End If

            Next

            withinTimeFrame = tmpListe

        End Get
    End Property

    Public Sub clearPhases()

        Try
            AllPhases.Clear()
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

    End Sub

    

    Public Sub keepPhase1consistent(ByVal phasenEnde As Integer)


        If Me.getPhase(1).dauerInDays < phasenEnde Then
            Me.getPhase(1).changeStartandDauerPhase1(0, phasenEnde)
            ' im Nebeneffekt wird ausserdem _Dauer aktualisiert  
            Dim projektLaengeInDays As Integer = Me.dauerInDays
        End If





    End Sub


    Public Sub clearBewertungen()
        Dim cPhase As clsPhase


        For p = 1 To Me.CountPhases
            cPhase = Me.getPhase(p)
            For r = 1 To cPhase.CountResults
                With cPhase.getResult(r)
                    .clearBewertungen()
                End With
            Next
        Next

    End Sub

    Public ReadOnly Property risikoKostenfaktor As Double
        Get
            Dim tmp As Double
            'tmp = (Me.Risiko - weightStrategicFit * Me.StrategicFit) / 100'
            ' wieso soll das Risiko geringer sein, wenn die Strategische Relevanz höher ist ? 
            tmp = Me.Risiko / 100
            If tmp < 0 Then
                tmp = 0
            End If
            risikoKostenfaktor = tmp
        End Get
    End Property
    ''' <summary>
    ''' kopiert die Attribute eines Projektes in newproject;  bei der Quelle handelt es sich um 
    ''' ein normales Projekt 
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <remarks></remarks>
    Public Overrides Sub copyAttrTo(ByRef newproject As clsProjekt)

        With newproject

            .farbe = Me.farbe
            .Schrift = Me.Schrift
            .Schriftfarbe = Me.Schriftfarbe
            .VorlagenName = Me.VorlagenName
            .Risiko = Me.Risiko
            .StrategicFit = Me.StrategicFit
            .Erloes = Me.Erloes
            .description = Me.description
            .volume = Me.volume
            .complexity = Me.complexity
            .businessUnit = Me.businessUnit
            .StartOffset = _StartOffset
            .startDate = _startDate
            .earliestStartDate = _earliestStartDate
            .latestStartDate = _latestStartDate
            .earliestStart = _earliestStart
            .latestStart = _latestStart
            .leadPerson = _leadPerson
            '.ProjectMarge = imarge
            .Status = _Status

        End With


    End Sub

    ''' <summary>
    ''' gibt die Bedarfe (Phasen / Rollen / Kostenarten / Ergebnis pro Monat zurück 
    ''' </summary>
    ''' <param name="mycollection">ist eine Liste mit Namen der zu betrachtenden Phasen-, Rollen-, Kosten bzw. Ergebnisse </param>
    ''' <param name="type">gibt an , worum es sich handelt; Phase, Rolle, Kostenart, Ergebnis</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBedarfeInMonths(mycollection As Collection, type As String) As Double()
        Get
            Dim i As Integer, k As Integer, projektDauer As Integer = Me.Dauer
            Dim valueArray() As Double
            Dim tempArray() As Double
            Dim riskarray() As Double
            Dim itemName As String
            Dim projektMarge As Double = Me.ProjectMarge

            If mycollection.Count = 0 Then
                Throw New ArgumentException("interner Fehler in getBedarfeinMonths: myCollection ist leer ")
            Else
                If projektDauer > 0 Then

                    ReDim valueArray(projektDauer - 1)
                    ReDim tempArray(projektDauer - 1)
                    ReDim riskarray(projektDauer - 1)

                    Select Case type
                        Case DiagrammTypen(0)

                            itemName = mycollection.Item(1)
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getPhasenBedarf(itemName)

                            For i = 2 To mycollection.Count
                                itemName = mycollection.Item(i)
                                tempArray = Me.getPhasenBedarf(itemName)
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) + tempArray(k)
                                Next
                            Next

                        Case DiagrammTypen(1)

                            itemName = mycollection.Item(1)
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getRessourcenBedarf(itemName)

                            For i = 2 To mycollection.Count
                                itemName = mycollection.Item(i)
                                tempArray = Me.getRessourcenBedarf(itemName)
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) + tempArray(k)
                                Next
                            Next

                        Case DiagrammTypen(2)

                            itemName = mycollection.Item(1)
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getKostenBedarf(itemName)


                            For i = 2 To mycollection.Count
                                itemName = mycollection.Item(i)
                                tempArray = Me.getKostenBedarf(itemName)
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) + tempArray(k)
                                Next
                            Next

                        Case DiagrammTypen(4)
                            Dim riskShare As Double
                            itemName = mycollection.Item(1)
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getGesamtKostenBedarf

                            If itemName = ergebnisChartName(0) Then
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) * projektMarge
                                Next

                            ElseIf itemName = ergebnisChartName(1) Then
                                riskShare = (Me.Risiko - weightStrategicFit * Me.StrategicFit) / 100
                                If riskShare < 0 Then
                                    riskShare = 0
                                End If

                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) * (projektMarge - riskShare)
                                Next

                            ElseIf itemName = ergebnisChartName(3) Then

                                riskShare = (Me.Risiko - weightStrategicFit * Me.StrategicFit) / 100
                                If riskShare < 0 Then
                                    riskShare = 0
                                End If

                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) * riskShare
                                Next
                            Else
                                Throw New ArgumentException("unbekannter Ergebnis-Typ ...")
                            End If

                        Case Else
                            Throw New ArgumentException("unbekannter Diagramm-Typ ...")

                    End Select
                Else
                    Throw New ArgumentException("Projekt " & Me.name & " hat keine Dauer ...")
                End If
            End If

            getBedarfeInMonths = valueArray

        End Get
    End Property

    ''' <summary>
    ''' gibt die Zahl der grünen/gelben/roten/grauen Bewertungen der Vergangenheit, der Zukunft oder beides an 
    ''' </summary>
    ''' <param name="colorIndex">0: keine Bewertung 1: grün 2: gelb 3: rot</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getNrColorIndexes(ByVal colorIndex As Integer) As Integer()
        Get
            Dim resultValues() As Integer
            Dim anzResults As Integer
            Dim anzPhasen As Integer
            Dim p As Integer, r As Integer
            Dim phase As clsPhase
            Dim result As clsResult
            Dim phasenStart As Integer, phasenEnde As Integer
            Dim monatsIndex As Integer


            If Me.Dauer > 0 Then

                ReDim resultValues(Me.Dauer - 1)


                'anzPhasen = Me.AllPhases.Count
                anzPhasen = MyBase.CountPhases

                For p = 1 To anzPhasen
                    phase = MyBase.getPhase(p)
                    With phase
                        ' Off1
                        anzResults = .CountResults
                        phasenStart = .relStart - 1
                        phasenEnde = .relEnde - 1


                        For r = 1 To anzResults

                            Try
                                result = .getResult(r)
                                monatsIndex = DateDiff(DateInterval.Month, Me.startDate, result.getDate)

                                ' Sicherstellen, daß Ergebnisse, die vor oder auch nach dem Projekt erreicht werden sollen, richtig behandelt werden 

                                If monatsIndex < 0 Then
                                    monatsIndex = 0
                                ElseIf monatsIndex > Me.Dauer - 1 Then
                                    monatsIndex = Me.Dauer - 1
                                End If

                                ' hier muss noch unterschieden werden, ob der ColorIndex = 0 ist: dann muss auch mitgezählt werden, wenn ein Result ohne Bewertung da ist ...

                                If result.getBewertung(1).colorIndex = colorIndex Then
                                    resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                End If

                                'Try
                                '    If result.getBewertung(1).colorIndex = colorIndex Then
                                '        resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                '    End If
                                'Catch ex1 As Exception
                                '    ' hierher kommt er, wenn es ein Result, aber keine Bewertung gibt 
                                '    If colorIndex = 0 Then
                                '        resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                '    End If
                                'End Try



                            Catch ex As Exception

                            End Try



                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen



            Else

                ReDim resultValues(0)
                resultValues(0) = ""

            End If

            getNrColorIndexes = resultValues

        End Get
    End Property


    Public ReadOnly Property getAllResults() As String()

        Get


            Dim ResultValues() As String
            Dim anzResults As Integer
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim result As clsResult
            Dim monatsIndex As Integer


            If Me.Dauer > 0 Then

                ReDim ResultValues(Me.Dauer - 1)
                For i = 0 To Me.Dauer - 1
                    ResultValues(i) = ""
                Next

                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzResults = .CountResults


                        For r = 1 To anzResults

                            result = .getResult(r)
                            monatsIndex = DateDiff(DateInterval.Month, Me.startDate, result.getDate)
                            ' Sicherstellen, daß Ergebnisse, die vor oder auch nach dem Projekt erreicht werden sollen, richtig behandelt werden 

                            If monatsIndex >= 0 And monatsIndex <= Me.Dauer - 1 Then

                                ResultValues(monatsIndex) = ResultValues(monatsIndex) & vbLf & result.name & _
                                                        " (" & result.getDate.ToShortDateString & ")"

                            End If





                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen



            Else

                ReDim ResultValues(0)
                ResultValues(0) = ""

            End If

            getAllResults = ResultValues

        End Get

    End Property


    ''' <summary>
    ''' gibt den Bedarf der Rolle in dem Monat X an; X=1 entspricht StartofCalendar usw.
    '''   
    ''' </summary>
    ''' <param name="mycollection"></param>
    ''' <param name="type"></param>
    ''' <param name="monat"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBedarfeInMonth(mycollection As Collection, type As String, monat As Integer) As Double


        Get
            Dim valueArray() As Double
            Dim tmpValue As Double
            Dim projektDauer As Integer = Me.Dauer
            Dim start As Integer = Me.Start

            If mycollection.Count = 0 Then
                Throw New ArgumentException("interner Fehler in getBedarfeinMonth: myCollection ist leer ")
            Else
                If projektDauer > 0 Then
                    ReDim valueArray(projektDauer - 1)
                    valueArray = Me.getBedarfeInMonths(mycollection, type)
                    If monat >= start And monat <= start + projektDauer - 1 Then
                        tmpValue = valueArray(monat - start)
                    Else
                        tmpValue = 0.0
                    End If
                Else
                    Throw New ArgumentException("getBedarfeinMonth: Projekt hat keine Dauer, " & Me.name)
                End If

            End If

            getBedarfeInMonth = tmpValue

        End Get
    End Property

    Public ReadOnly Property hasDifferentRoleNeeds(ByVal compareProj As clsProjekt, roleName As String) As Boolean
        Get
            Dim myArray() As Double
            Dim hisArray() As Double
            Dim istIdentisch As Boolean = True
            Dim i As Integer


            Try
                myArray = Me.getRessourcenBedarf(roleName)
                hisArray = compareProj.getRessourcenBedarf(roleName)
                If myArray.Length <> hisArray.Length Then
                    istIdentisch = False
                End If
                i = 0
                While i <= myArray.Length - 1 And istIdentisch
                    If myArray(i) <> hisArray(i) Then
                        istIdentisch = False
                    Else
                        i = i + 1
                    End If
                End While
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            hasDifferentRoleNeeds = Not istIdentisch

        End Get
    End Property

    Public ReadOnly Property hasDifferentCostNeeds(ByVal compareProj As clsProjekt, costName As String) As Boolean
        Get
            Dim myArray() As Double
            Dim hisArray() As Double
            Dim istIdentisch As Boolean = True
            Dim i As Integer

            Try
                myArray = Me.getKostenBedarf(costName)
                hisArray = compareProj.getKostenBedarf(costName)
                If myArray.Length <> hisArray.Length Then
                    istIdentisch = False
                End If
                i = 0
                While i <= myArray.Length - 1 And istIdentisch
                    If myArray(i) <> hisArray(i) Then
                        istIdentisch = False
                    Else
                        i = i + 1
                    End If
                End While

            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            hasDifferentCostNeeds = Not istIdentisch

        End Get
    End Property

    ''' <summary>
    ''' kopiert alle Meilensteine, aber ohne Bewertung 
    ''' </summary>
    ''' <param name="newproj"></param>
    ''' <remarks></remarks>
    Public Sub copyResultsTo(ByRef newproj As clsProjekt)

        Dim newresult As clsResult
        Dim newphase As clsPhase

        ' Kopiere die Ampel - und die Ampel-Bewertung
        With newproj
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung
        End With

        For Each cphase In MyBase.Liste

            Try
                newphase = newproj.getPhase(cphase.name)
                ' wenn gefunden dann alle Results kopieren 
                For r = 1 To cphase.CountResults
                    newresult = New clsResult(parent:=newphase)
                    cphase.getResult(r).CopyToWithoutBewertung(newresult)
                    newphase.addresult(newresult)
                Next

            Catch ex As Exception
                ' in diesem Falle gibt es die komplette Phase in dem Projekt nicht mehr 
                ' dann muss auch nichts gemacht werden 
            End Try


        Next

    End Sub



    Public Sub copyBewertungenTo(ByRef newproj As clsProjekt)

        Dim newresult As clsResult
        Dim newphase As clsPhase

        ' Kopiere die Ampel - und die Ampel-Bewertung
        With newproj
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung
        End With

        For Each cphase In MyBase.Liste

            Try
                newphase = newproj.getPhase(cphase.name)
                ' wenn gefunden dann alle Results kopieren 
                For r = 1 To cphase.CountResults
                    newresult = New clsResult(parent:=newphase)
                    cphase.getResult(r).CopyTo(newresult)
                    newphase.addresult(newresult)
                Next

            Catch ex As Exception
                ' in diesem Falle gibt es die komplette Phase in dem Projekt nicht mehr 
                ' dann muss auch nichts gemacht werden 
            End Try


        Next

    End Sub


    Public Overrides Sub CopyTo(ByRef newproject As clsProjekt)

        Dim newphase As clsPhase

        Call copyAttrTo(newproject)

        For Each hphase In MyBase.Liste
            newphase = New clsPhase(newproject)
            hphase.CopyTo(newphase)
            newproject.AddPhase(newphase)
        Next


    End Sub


    Public Overrides Sub korrCopyTo(ByRef newproject As clsProjekt, ByVal startdate As Date, ByVal endedate As Date)
        Dim p As Integer
        Dim newphase As clsPhase
        Dim oldphase As clsPhase
        Dim ProjectDauerInDays As Integer
        Dim CorrectFactor As Double

        Call copyAttrTo(newproject)

        With newproject
            .startDate = startdate
            .earliestStart = _earliestStart
            .latestStart = _latestStart

            ProjectDauerInDays = calcDauerIndays(startdate, endedate)
            CorrectFactor = ProjectDauerInDays / Me.dauerInDays


            For p = 1 To Me.CountPhases

                oldphase = Me.getPhase(p)
                newphase = New clsPhase(newproject)

                oldphase.korrCopyTo(newphase, CorrectFactor)

                .AddPhase(newphase)

            Next p


        End With

    End Sub

    ''' <summary>
    ''' gibt zurück, 
    ''' in gettimeCostColor(0): ob das Projekt schneller oder langsamer als das Vergleichsprojekt ist ;
    ''' in gettimeCostColor(1): ob das Projekt günstiger oder teurer als das Vergleichsprojekt ist ;
    ''' in gettimeCostColor(2): welche Bewertung der nächste bzw. letzte  Meilenstein (in Abhängigkeit von Auswahl) hat 
    ''' 
    ''' </summary>
    ''' <param name="vproj"></param>
    ''' meist der Planungs-Stand zur Zeit der Beauftragung, oder aber der letzte Stand
    ''' <param name="auswahl">
    ''' 0: Vergleiche Projektende
    ''' 1: vergleiche nächsten Meilenstein 
    ''' </param>
    ''' <param name="refDate">
    ''' bestimmt das Datum, ab dem der nächstgelegene nächste Meilenstein gesucht wird</param>
    ''' <value>
    ''' gibt die Time Kennzahl zurück: "kleiner 1": ist schneller; "größer 1": ist langsamer
    ''' </value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTimeCostColor(ByVal vproj As clsProjekt, ByVal auswahl As Integer, ByVal showAbsoluteDiff As Boolean, ByVal refDate As Date) As Double()
        Get
            Dim tmpValue(2) As Double
            Dim curMsName As String = ""
            Dim curPhName As String = ""
            Dim curAbstand As Integer = 10000
            Dim tmpAbstand As Integer
            Dim tmpPhase As clsPhase
            Dim tmpColor As Integer = -1
            Dim anzResults As Integer
            Dim relMonat1 As Integer, relMonat2 As Integer
            Dim chkDate1 As Date, chkDate2 As Date, tmpDate As Date

            If auswahl = 0 Then
                If showAbsoluteDiff Then
                    tmpValue(0) = Me.dauerInDays - vproj.dauerInDays
                    tmpValue(1) = Me.getSummeKosten - vproj.getSummeKosten
                    tmpValue(2) = Me.ampelStatus
                Else
                    tmpValue(0) = Me.dauerInDays / vproj.dauerInDays
                    tmpValue(1) = Me.getSummeKosten / vproj.getSummeKosten
                    tmpValue(2) = Me.ampelStatus
                End If


            ElseIf auswahl = 1 Then

                Dim nullWert As Integer = DateDiff(DateInterval.Day, Me.startDate, refDate) + 1

                If nullWert > Me.dauerInDays Then
                    ' Projekt ist bereits beendet ...
                    If showAbsoluteDiff Then
                        tmpValue(0) = Me.dauerInDays - vproj.dauerInDays
                        tmpValue(1) = Me.getSummeKosten - vproj.getSummeKosten
                        tmpValue(2) = Me.ampelStatus
                    Else
                        tmpValue(0) = Me.dauerInDays / vproj.dauerInDays
                        tmpValue(1) = Me.getSummeKosten / vproj.getSummeKosten
                        tmpValue(2) = Me.ampelStatus
                    End If
                Else

                    Dim vglWert1 As Integer = -1
                    Dim vglWert2 As Integer = -1

                    ' bestimme die Phase und Meilenstein , der als nächstes nach refdate kommt 
                    For p = 1 To Me.CountPhases

                        tmpPhase = Me.getPhase(p)
                        anzResults = tmpPhase.CountResults


                        For r = 1 To anzResults
                            tmpDate = tmpPhase.getResult(r).getDate
                            tmpAbstand = DateDiff(DateInterval.Day, refDate, tmpDate)
                            If tmpAbstand > 0 And tmpAbstand < curAbstand Then
                                curMsName = tmpPhase.getResult(r).name
                                curPhName = tmpPhase.name
                                curAbstand = tmpAbstand
                                chkDate1 = tmpDate
                                tmpColor = tmpPhase.getResult(r).getBewertung(1).colorIndex
                            End If
                        Next

                        tmpDate = tmpPhase.getEndDate
                        ' falls es in dieser Phase keinen Meilenstein gab ... oder falls das Phasen Ende noch vor dem Meilenstein lag
                        If tmpPhase.dauerInDays > nullWert And tmpPhase.dauerInDays - nullWert < curAbstand Then
                            curMsName = ""
                            curPhName = tmpPhase.name
                            curAbstand = tmpPhase.dauerInDays - nullWert
                            chkDate1 = tmpDate
                            If tmpColor = -1 Then
                                tmpColor = Me.ampelStatus
                            End If
                        End If

                    Next

                    ' jetzt ist sichergestellt , daß es zumindest curPhName (current PhaseName) gibt, evtl auch curMsName (current MilestoneName)
                    If curPhName <> "" Then
                        vglWert1 = curAbstand + nullWert
                        ' jetzt muss der Vergleichswert2 bestimmt werden ...
                        tmpPhase = vproj.getPhase(curPhName)

                        If IsNothing(tmpPhase) Then
                            ' im vergleichsprojekt gibt es die Phase gar nicht , also muss auf das Gesamtprojekt verglichen werden 
                            vglWert1 = Me.dauerInDays
                            vglWert2 = vproj.dauerInDays
                            chkDate1 = Me.endeDate
                            chkDate2 = vproj.endeDate
                        Else

                            If curMsName <> "" Then
                                Dim tmpResult As clsResult
                                tmpResult = tmpPhase.getResult(curMsName)
                                ' gibt es den Meilenstein in der Phase ? 
                                If IsNothing(tmpResult) Then

                                    ' die beiden Phasen-Ende als Vergleichskriterien nehmen 
                                    With Me.getPhase(curPhName)
                                        vglWert1 = .startOffsetinDays + .dauerInDays
                                        chkDate1 = .getEndDate
                                    End With

                                    With tmpPhase
                                        vglWert2 = .startOffsetinDays + .dauerInDays
                                        chkDate2 = .getEndDate
                                    End With

                                Else

                                    With tmpPhase
                                        vglWert2 = .startOffsetinDays + tmpResult.offset
                                        chkDate2 = tmpResult.getDate
                                    End With

                                End If

                            Else
                                With Me.getPhase(curPhName)
                                    vglWert1 = .startOffsetinDays + .dauerInDays
                                    chkDate1 = .getEndDate
                                End With

                                With tmpPhase
                                    vglWert2 = .startOffsetinDays + .dauerInDays
                                    chkDate2 = .getEndDate
                                End With

                            End If

                        End If

                        relMonat1 = getColumnOfDate(chkDate1) - Me.Start
                        relMonat2 = getColumnOfDate(chkDate2) - vproj.Start

                        If showAbsoluteDiff Then
                            tmpValue(0) = vglWert1 - vglWert2

                            ' nun jeweils die Summen bis zum angegebenen Monat aufsummieren ....
                            ' ... und die Kennzahl berechnen 
                            tmpValue(1) = Me.getSummeKosten(relMonat1) - vproj.getSummeKosten(relMonat2)
                            tmpValue(2) = tmpColor
                        Else
                            tmpValue(0) = vglWert1 / vglWert2
                            ' nun jeweils die Summen bis zum angegebenen Monat aufsummieren ....
                            ' ... und die Kennzahl berechnen 
                            tmpValue(1) = Me.getSummeKosten(relMonat1) / vproj.getSummeKosten(relMonat2)
                            tmpValue(2) = tmpColor
                        End If



                    Else
                        If showAbsoluteDiff Then
                            tmpValue(0) = Me.dauerInDays - vproj.dauerInDays
                            tmpValue(1) = Me.getSummeKosten - vproj.getSummeKosten
                            tmpValue(2) = Me.ampelStatus
                        Else
                            tmpValue(0) = Me.dauerInDays / vproj.dauerInDays
                            tmpValue(1) = Me.getSummeKosten / vproj.getSummeKosten
                            tmpValue(2) = Me.ampelStatus
                        End If
                    End If

                End If



            Else
                ' Fehler: Auswahl nicht definiert 
                Throw New Exception("Fehler in getTimeIndex")
            End If


            ' Sicherstellen Konsistenzbedingung: Farbe kann nicht negativ sein  
            If tmpValue(2) < 0 Then
                tmpValue(2) = 0
            End If

            getTimeCostColor = tmpValue


        End Get
    End Property

    Public ReadOnly Property getTimeTimeColor(ByVal vproj As clsProjekt, ByVal showAbsoluteDiff As Boolean, ByVal refDate As Date) As Double()
        Get
            Dim tmpValue(2) As Double
            Dim curMsName As String = ""
            Dim curPhName As String = ""
            Dim curAbstand As Integer = 10000
            Dim tmpAbstand As Integer
            Dim tmpPhase As clsPhase
            Dim tmpColor As Integer = -1
            Dim anzResults As Integer
            Dim chkDate1 As Date, chkDate2 As Date, tmpDate As Date

            ' hier ist jetzt klar, was die Unterschiede im Vergleich Projektende/Projektende sind
            If showAbsoluteDiff Then
                tmpValue(0) = Me.dauerInDays - vproj.dauerInDays

            Else
                tmpValue(0) = Me.dauerInDays / vproj.dauerInDays

            End If



            Dim nullWert As Integer = DateDiff(DateInterval.Day, Me.startDate, refDate) + 1

            If nullWert > Me.dauerInDays Then
                ' Projekt ist bereits beendet ...
                If showAbsoluteDiff Then
                    tmpValue(1) = tmpValue(0)
                    tmpValue(2) = Me.ampelStatus
                Else
                    tmpValue(1) = tmpValue(0)
                    tmpValue(2) = Me.ampelStatus
                End If
            Else

                Dim vglWert1 As Integer = -1
                Dim vglWert2 As Integer = -1

                ' bestimme die Phase und Meilenstein , der als nächstes nach refdate kommt 
                For p = 1 To Me.CountPhases

                    tmpPhase = Me.getPhase(p)
                    anzResults = tmpPhase.CountResults


                    For r = 1 To anzResults
                        tmpDate = tmpPhase.getResult(r).getDate
                        tmpAbstand = DateDiff(DateInterval.Day, refDate, tmpDate)
                        If tmpAbstand > 0 And tmpAbstand < curAbstand Then
                            curMsName = tmpPhase.getResult(r).name
                            curPhName = tmpPhase.name
                            curAbstand = tmpAbstand
                            chkDate1 = tmpDate
                            tmpColor = tmpPhase.getResult(r).getBewertung(1).colorIndex
                        End If
                    Next

                    tmpDate = tmpPhase.getEndDate
                    ' falls es in dieser Phase keinen Meilenstein gab ... oder falls das Phasen Ende noch vor dem Meilenstein lag
                    If tmpPhase.dauerInDays > nullWert And tmpPhase.dauerInDays - nullWert < curAbstand Then
                        curMsName = ""
                        curPhName = tmpPhase.name
                        curAbstand = tmpPhase.dauerInDays - nullWert
                        chkDate1 = tmpDate
                        If tmpColor = -1 Then
                            tmpColor = Me.ampelStatus
                        End If
                    End If

                Next

                ' jetzt ist sichergestellt , daß es zumindest curPhName (current PhaseName) gibt, evtl auch curMsName (current MilestoneName)
                If curPhName <> "" Then
                    vglWert1 = curAbstand + nullWert
                    ' jetzt muss der Vergleichswert2 bestimmt werden ...
                    tmpPhase = vproj.getPhase(curPhName)

                    If IsNothing(tmpPhase) Then
                        ' im vergleichsprojekt gibt es die Phase gar nicht , also muss auf das Gesamtprojekt verglichen werden 
                        vglWert1 = Me.dauerInDays
                        vglWert2 = vproj.dauerInDays
                        chkDate1 = Me.endeDate
                        chkDate2 = vproj.endeDate
                    Else

                        If curMsName <> "" Then
                            Dim tmpResult As clsResult
                            tmpResult = tmpPhase.getResult(curMsName)
                            ' gibt es den Meilenstein in der Phase ? 
                            If IsNothing(tmpResult) Then

                                ' die beiden Phasen-Ende als Vergleichskriterien nehmen 
                                With Me.getPhase(curPhName)
                                    vglWert1 = .startOffsetinDays + .dauerInDays
                                    chkDate1 = .getEndDate
                                End With

                                With tmpPhase
                                    vglWert2 = .startOffsetinDays + .dauerInDays
                                    chkDate2 = .getEndDate
                                End With

                            Else

                                With tmpPhase
                                    vglWert2 = .startOffsetinDays + tmpResult.offset
                                    chkDate2 = tmpResult.getDate
                                End With

                            End If

                        Else
                            With Me.getPhase(curPhName)
                                vglWert1 = .startOffsetinDays + .dauerInDays
                                chkDate1 = .getEndDate
                            End With

                            With tmpPhase
                                vglWert2 = .startOffsetinDays + .dauerInDays
                                chkDate2 = .getEndDate
                            End With

                        End If

                    End If


                    If showAbsoluteDiff Then
                        tmpValue(1) = vglWert1 - vglWert2
                        tmpValue(2) = tmpColor
                    Else

                        tmpValue(1) = vglWert1 / vglWert2
                        tmpValue(2) = tmpColor
                    End If



                Else
                    If showAbsoluteDiff Then
                        tmpValue(1) = Me.dauerInDays - vproj.dauerInDays
                        tmpValue(2) = Me.ampelStatus
                    Else
                        tmpValue(1) = Me.dauerInDays / vproj.dauerInDays
                        tmpValue(2) = Me.ampelStatus
                    End If
                End If

            End If




            ' Sicherstellen Konsistenzbedingung: Farbe kann nicht negativ sein  
            If tmpValue(2) < 0 Then
                tmpValue(2) = 0
            End If

            getTimeTimeColor = tmpValue


        End Get
    End Property

    '
    ' übergibt in Project Marge die berechnete Marge: Erloes - Kosten
    '
    Public ReadOnly Property ProjectMarge() As Double


        Get
            Dim gk As Double = 10.0
            Try
                gk = Me.getGesamtKostenBedarf.Sum
                ' prüfen , ob die Marge konsistent ist mit Verhältnis Erlös und Kosten  ... 

                If gk > 0 Then
                    ProjectMarge = (Me.Erloes - gk) / gk
                Else
                    ProjectMarge = 0
                End If

            Catch ex As Exception
                'Call MsgBox("Projekt: " & Me.name & vbLf & "gk: " & gk.ToString)
                ProjectMarge = 0
            End Try


        End Get

        'Set(value As Double)

        '    imarge = value

        'End Set

    End Property

    Public Overrides Property earliestStart() As Integer

        Get
            earliestStart = _earliestStart
        End Get

        Set(value As Integer)
            Dim heuteColumn As Integer = getColumnOfDate(Date.Today)
            Dim reasonableValue As Integer

            If value <= 0 Then
                If Me.Start + value > heuteColumn Then
                    ' es ist zugelassen 
                    _earliestStart = value
                Else
                    ' das Projekt kann frühestens im Folge Monat beginnen  
                    reasonableValue = heuteColumn + 1 - Me.Start
                    If reasonableValue > 0 Then
                        reasonableValue = 0
                    End If
                    _earliestStart = reasonableValue
                End If

            End If
        End Set
    End Property


    Public Overrides Property latestStart() As Integer

        Get
            latestStart = _latestStart
        End Get

        Set(value As Integer)
            If value > 0 Then
                _latestStart = value
            End If

        End Set

    End Property

    Public ReadOnly Property Start() As Integer

        Get
            Start = _Start
        End Get


    End Property

    Public Property Status() As String
        Get
            Status = _Status
        End Get
        Set(value As String)
            If value = ProjektStatus(0) Then
                _Status = value
            ElseIf value = ProjektStatus(1) Or value = ProjektStatus(2) Or _
                                               value = ProjektStatus(3) Or _
                                               value = ProjektStatus(4) Then
                _Status = value
                ' 2.5.2014 ur: Die nächsten Befehle sind auskommentiert, weil ein beauftragtes Projekt
                ' nicht zwangsweise bereits gestartet wurde 
                '_earliestStart = 0
                '_latestStart = 0
                '_earliestStartDate = _startDate
                '_latestStartDate = _startDate
            Else
                Call MsgBox("unzulässiger Wert für Status")
            End If
        End Set
    End Property

    Public Property StartOffset As Integer
        Get
            StartOffset = _StartOffset
        End Get

        Set(value As Integer)
            If value >= _earliestStart And value <= _latestStart Then
                _StartOffset = value
            Else
                Call MsgBox("unzulässiger Wert für StartOffset ...")
            End If
        End Set

    End Property

    ''' <summary>
    ''' errechnet  die Position (top, left) und Größe (width, height) des Shapes, das das Projekt repräsentieren soll 
    ''' Voraussetzung: tfzeile und tfspalte haben den für das Projekt richtigen Wert  
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="Left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub CalculateShapeCoord(ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        Dim projektStartdate As Date = Me.startDate
        Dim startpunkt As Integer = DateDiff(DateInterval.Day, StartofCalendar, projektStartdate)

        If startpunkt < 0 Then
            Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
        End If

        Dim projektlaenge As Integer = Me.dauerInDays

        If Me.tfZeile <= 1 Then
            Me.tfZeile = 2
        End If

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.Dauer > 0 Then
            top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight

            ' neue Positionierung: Tagesgenau 
            left = (startpunkt / 365) * boxWidth * 12

            ' check it, notfalls korrigieren ... 
            If System.Math.Truncate(left / boxWidth) + 1 < Me.Start Then

                Do Until System.Math.Truncate(left / boxWidth) + 1 = Me.Start
                    left = left + 1
                Loop

            ElseIf System.Math.Truncate(left / boxWidth) + 1 > Me.Start Then

                Do Until System.Math.Truncate(left / boxWidth) + 1 = Me.Start
                    left = left - 1
                Loop

            End If

            width = ((projektlaenge) / 365) * boxWidth * 12

            ' Alte Positionierung: auf Monat 
            'left = (Me.tfspalte - 1) * boxWidth + 0.5
            'width = Me.Dauer * boxWidth - 1
            height = 0.8 * boxHeight
        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub

    ''' <summary>
    ''' berechnet die Koordinaten der Phase mit Nummer  phaseNr. 
    ''' </summary>
    ''' <param name="phaseNr"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub CalculateShapeCoord(ByVal phaseNr As Integer, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)

        Dim cphase As clsPhase

        Try

            Dim projektStartdate As Date = Me.startDate
            Dim startpunkt As Integer = DateDiff(DateInterval.Day, StartofCalendar, projektStartdate)

            If startpunkt < 0 Then
                Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
            End If

            cphase = Me.getPhase(phaseNr)
            Dim phasenStart As Integer = startpunkt + cphase.startOffsetinDays
            Dim phasenDauer As Integer = cphase.dauerInDays



            If Me.tfZeile > 1 And phasenStart >= 1 And phasenDauer > 0 Then

                If phaseNr = 1 Then
                    top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight
                    ' Änderung 28.11 jetzt wird tagesgenau positioniert 
                    left = (phasenStart / 365) * boxWidth * 12
                    width = ((phasenDauer) / 365) * boxWidth * 12

                    ' Alte Positionierung, als nur auf den Monat positioniert wurde 
                    'left = (phasenStart - 1) * boxWidth + 0.5
                    'width = phasenDauer * boxWidth - 1
                    height = 0.8 * boxHeight
                Else
                    If top <= 0 Then
                        top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + 0.1 * boxHeight
                    Else
                        ' nichts tun : top wird an der Aufrufenden Stelle gesetzt
                        ' zeichneProjektinPlantafel2 Änderung 18.3.14 
                    End If

                    left = (phasenStart / 365) * boxWidth * 12
                    width = ((phasenDauer) / 365) * boxWidth * 12

                    ' Alte Positionierung, als nur auf dne Monat positioniert wurde 
                    'left = (phasenStart - 1) * boxWidth + 0.5
                    'width = phasenDauer * boxWidth - 1
                    height = 0.6 * boxHeight
                End If


            Else
                Throw New ArgumentException("es kann kein Shape berechnet werden für : " & cphase.name)
            End If

        Catch ex As Exception
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name & "Phase: " & phaseNr.ToString)
        End Try


    End Sub

    ''' <summary>
    ''' gibt für die angegebene Phasen-Nummer den zeilenoffset zurück sowie die 
    ''' Werte für top, left, width, height des Phasen-Shapes
    ''' </summary>
    ''' <param name="phaseNr"></param>
    ''' <param name="zeilenOffset"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub CalculateShapeCoord(ByVal phaseNr As Integer, ByRef zeilenOffset As Integer,
                                       ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)
        Dim cphase As clsPhase
        Dim phasenName As String
        Dim lastEndDate As Date = StartofCalendar.AddDays(-1)


        If phaseNr > Me.CountPhases Then
            Throw New ArgumentException("es gibt diese Phasen-Numer nicht: " & phaseNr & vbLf & _
                                         "Projekt: " & Me.name & ", Anzahl Phasen: " & Me.CountPhases)
        End If

        For i = 1 To phaseNr

            With Me.getPhase(i)

                phasenName = .name
                If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
                    zeilenOffset = zeilenOffset + 1
                    lastEndDate = StartofCalendar.AddDays(-1)
                End If

                If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
                    lastEndDate = .getEndDate
                End If

            End With
        Next


        Try

            Dim projektStartdate As Date = Me.startDate
            Dim startpunkt As Integer = DateDiff(DateInterval.Day, StartofCalendar, projektStartdate)

            If startpunkt < 0 Then
                Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
            End If

            cphase = Me.getPhase(phaseNr)
            Dim phasenStart As Integer = startpunkt + cphase.startOffsetinDays
            Dim phasenDauer As Integer = cphase.dauerInDays



            If Me.tfZeile > 1 And phasenStart >= 1 And phasenDauer > 0 Then

                ' Änderung 18.3.14 Zeilenoffset gibt an, in die wievielte Zeile das geschrieben werden soll 
                If phaseNr = 1 Then
                    top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + 0.1 * boxHeight
                    left = (phasenStart / 365) * boxWidth * 12
                    width = ((phasenDauer) / 365) * boxWidth * 12
                    height = 0.8 * boxHeight
                Else
                    'top = topOfMagicBoard + (Me.tfZeile + zeilenOffset - 1) * boxHeight + 0.1 * boxHeight
                    top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + 0.1 * boxHeight + (zeilenOffset + 1) * boxHeight * 0.5
                    left = (phasenStart / 365) * boxWidth * 12
                    width = ((phasenDauer) / 365) * boxWidth * 12
                    height = 0.33 * boxHeight
                End If


            Else
                Throw New ArgumentException("es kann kein Shape berechnet werden für : " & cphase.name)
            End If

        Catch ex As Exception
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name & "Phase: " & phaseNr.ToString)
        End Try


    End Sub


    Public Sub calculateResultCoord(ByVal resultDate As Date, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)



        Dim msStart As Integer = DateDiff(DateInterval.Day, StartofCalendar, resultDate)
        Dim faktor As Double = 1.0

        'Dim tagebisResult As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start - 1), resultDate)
        'Dim ratio As Double = tagebisResult / anzahlTage

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.Dauer > 0 Then
            top = topOfMagicBoard + (Me.tfZeile - 1.0) * boxHeight - boxWidth / 2
            left = (msStart / 365) * boxWidth * 12 - boxWidth * 0.5 * faktor
            width = boxWidth
            height = boxWidth
        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub

    Public Sub calculateResultCoord(ByVal resultDate As Date, ByVal zeilenOffset As Integer, _
                                    ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        'Dim endDatum As Date = StartofCalendar.AddMonths(Me.Start - 1 + Dauer).AddDays(-1)
        Dim diffMonths As Integer = DateDiff(DateInterval.Month, StartofCalendar, resultDate)
        Dim dayOfResult As Integer = resultDate.Day
        Dim msStart As Integer = DateDiff(DateInterval.Day, StartofCalendar, resultDate)

        Dim faktor As Double = 0.66

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.Dauer > 0 Then

            ' Änderung 18.3.14 Zeilenoffset gibt an, in die wievielte Zeile das geschrieben werden soll 
            If zeilenOffset = 0 Then
                top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight - boxWidth * 0.5 * faktor
                left = (msStart / 365) * boxWidth * 12 - boxWidth * 0.5 * faktor
                width = faktor * boxWidth
                height = faktor * boxWidth
            Else
                'top = topOfMagicBoard + (Me.tfZeile + zeilenOffset - 1) * boxHeight + 0.1 * boxHeight
                top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + (zeilenOffset + 1) * boxHeight * 0.5 - boxWidth * 0.5 * faktor
                left = (msStart / 365) * boxWidth * 12 - boxWidth * 0.5 * faktor
                width = faktor * boxWidth
                height = faktor * boxWidth
            End If


        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub

    Public Sub calculateRoundedKPI(ByRef budget As Double, ByRef personalKosten As Double, ByRef sonstKosten As Double, ByRef risikoKosten As Double, ByRef ergebnis As Double)

        With Me
            Dim gk As Double = .getSummeKosten

            budget = System.Math.Round(.Erloes, mode:=MidpointRounding.ToEven)

            risikoKosten = System.Math.Round(.risikoKostenfaktor * gk, mode:=MidpointRounding.ToEven)

            personalKosten = System.Math.Round(.getAllPersonalKosten.Sum, mode:=MidpointRounding.ToEven)

            sonstKosten = System.Math.Round(.getGesamtAndereKosten.Sum, mode:=MidpointRounding.ToEven)

            ergebnis = budget - (risikoKosten + personalKosten + sonstKosten)

        End With

    End Sub



    Public Sub calculateStatusCoord(ByVal resultDate As Date, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        ' es wird geprüft, ob das Projekt nicht schon beendet ist oder noch gar nicht angefangen hat 
        Dim endDatum As Date = Me.startDate.AddDays(Me.dauerInDays - 1)

        If DateDiff(DateInterval.Month, Me.startDate, resultDate) < 0 Then
            ' Projekt-Start hat noch gar nicht stattgefunden 
            resultDate = Me.startDate
        ElseIf DateDiff(DateInterval.Month, resultDate, endDatum) < 0 Then
            resultDate = endDatum
        End If



        Dim diffMonths As Integer = DateDiff(DateInterval.Month, StartofCalendar, resultDate)
        'Dim dayOfResult As Integer = resultDate.Day
        Dim dayOfResult As Integer = 15 ' wähle die Mitte des Monats

        'Dim tagebisResult As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start - 1), resultDate)
        'Dim ratio As Double = tagebisResult / anzahlTage

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.Dauer > 0 Then
            top = topOfMagicBoard + (Me.tfZeile - 1.0) * boxHeight
            left = diffMonths * boxWidth + dayOfResult * (boxWidth / 30) - 0.5 * boxWidth

            width = boxWidth
            height = boxWidth
        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub


    Public Sub New()

        AllPhases = New List(Of clsPhase)
        diffToPrev = False
        relStart = 1
        _leadPerson = ""
        iDauer = 0
        _StartOffset = 0
        _Start = 0
        _startDate = NullDatum
        _earliestStart = 0
        _latestStart = 0
        _Status = ProjektStatus(0)
        _shpUID = ""
        _variantName = ""   ' ur:25.6.2014: hinzugefügt, da sonst in der DB variantName mal "" und mal Nothing ist
        _timeStamp = Date.Now


    End Sub

    Public Sub New(ByVal projektStart As Integer, ByVal earliestValue As Integer, ByVal latestValue As Integer)

        AllPhases = New List(Of clsPhase)
        diffToPrev = False
        relStart = 1
        _leadPerson = ""
        iDauer = 0
        _StartOffset = 0

        _Start = projektStart
        _earliestStart = earliestValue
        _latestStart = latestValue

        _startDate = StartofCalendar.AddMonths(projektStart)
        _earliestStartDate = _startDate.AddMonths(_earliestStart)
        _latestStartDate = _startDate.AddMonths(_latestStart)

        _Status = ProjektStatus(0)
        _shpUID = ""
        _variantName = ""
        _timeStamp = Date.Now

    End Sub

    Public Sub New(ByVal startDate As Date, ByVal earliestStartdate As Date, ByVal latestStartdate As Date)

        AllPhases = New List(Of clsPhase)
        relStart = 1
        _leadPerson = ""
        iDauer = 0
        _StartOffset = 0

        _startDate = startDate
        _earliestStartDate = earliestStartdate
        _latestStartDate = latestStartdate

        _Start = DateDiff(DateInterval.Month, StartofCalendar, startDate) + 1
        _earliestStart = DateDiff(DateInterval.Month, startDate, earliestStartdate)
        _latestStart = DateDiff(DateInterval.Month, startDate, latestStartdate)

        _Status = ProjektStatus(0)
        _variantName = ""
        _timeStamp = Date.Now

    End Sub
End Class
