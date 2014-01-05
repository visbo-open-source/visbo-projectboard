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

    Private NullDatum As Date = "23.6.1914"



    ' Deklarationen der Events 
    Public Property shpUID As String
    Public Property name As String
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

                            phaseStart = Me.startDate.AddDays(phase.startOffsetinDays)
                            phaseEnd = Me.startDate.AddDays(phase.startOffsetinDays + phase.dauerInDays - 1)

                            ReDim phaseValues(phase.relEnde - phase.relStart)

                            With phase
                                For i = 0 To .relEnde - .relStart

                                    If i = 0 Then

                                        numberOfDays = Max(0.0, DateDiff(DateInterval.Day, phaseStart, StartofCalendar.AddMonths(Me.Start + .relStart - 1).AddDays(-1)))
                                        anteil = numberOfDays / 365 * 12
                                        phaseValues(i) = phaseValues(i) + Min(1.0, anteil)

                                    ElseIf i = .relEnde - .relStart Then

                                        numberOfDays = Max(0.0, DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start + .relEnde - 2), phaseEnd))
                                        anteil = numberOfDays / 365 * 12
                                        phaseValues(i) = phaseValues(i) + Min(1.0, anteil)

                                    Else

                                        phaseValues(i) = phaseValues(i) + 1

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

    Public Overrides ReadOnly Property Dauer() As Integer


        Get
            Dim i As Integer
            Dim max As Double = 0
            Dim maxM As Integer

            ' neue Bestimmung der Dauer 

            For i = 1 To AllPhases.Count

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays - 1 Then
                        max = .startOffsetinDays + .dauerInDays - 1
                    End If

                    For m = 1 To .CountResults
                        If max < .startOffsetinDays + .getResult(m).offset Then
                            max = .startOffsetinDays + .getResult(m).offset
                        End If
                    Next

                End With

            Next i

            maxM = DateDiff(DateInterval.Month, startDate, startDate.AddDays(max)) + 1


            If maxM <> _Dauer Then
                _Dauer = maxM
            End If


            Dauer = _Dauer


        End Get

    End Property


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


    Public Property startDate As Date
        Get
            startDate = _startDate
        End Get

        Set(value As Date)

            Dim olddate As Date = _startDate
            Dim differenzInTagen As Integer = DateDiff(DateInterval.Day, olddate, value)
            Dim updatePhases As Boolean = False

            If DateDiff(DateInterval.Month, _startDate, value) = 0 Then
                ' dann darf noch verändert werden, auch wenn es schon beauftragt wurde  
                _startDate = value
                _Start = DateDiff(DateInterval.Month, StartofCalendar, value) + 1
                updatePhases = True

            ElseIf _startDate = NullDatum Then
                _startDate = value
                _Start = DateDiff(DateInterval.Month, StartofCalendar, value) + 1
            Else
                ' es muss geprüft werden, ob es noch im Planungs-Stadium ist: nur dann darf noch verschoben werden ...
                If _Status = ProjektStatus(0) Then
                    _startDate = value
                    _Start = DateDiff(DateInterval.Month, StartofCalendar, value) + 1
                    updatePhases = True
                Else
                    Throw New ArgumentException("der Startzeitpunkt kann nicht mehr verändert werden ... ")
                End If

            End If

            'If updatePhases Then

            '    If olddate <> NullDatum And differenzInTagen <> 0 Then
            '        Dim chkValue As Integer, oldvalue As Integer
            '        ' jetzt müssen die Phasen-Werte verändert werden 
            '        For p = 1 To Me.CountPhases

            '        Next
            '    End If

            'End If


        End Set
    End Property


    Public Property earliestStartDate As Date
        Get
            earliestStartDate = _earliestStartDate
        End Get
        Set(value As Date)
            'Dim Heute As Date = Now

            _earliestStartDate = value

            'If _Status = ProjektStatus(1) Or _Status = ProjektStatus(2) Or _
            '                                 _Status = ProjektStatus(2) Then
            '    Throw New ArgumentException("der Startzeitpunkt kann nicht mehr verändert werden ... ")
            'Else
            '    If DateDiff(DateInterval.Month, StartofCalendar, value) + 1 > 0 And DateDiff(DateInterval.Month, _startDate, value) <= 0 Then
            '        If DateDiff(DateInterval.Month, Heute, value) > 0 Then
            '            _earliestStartDate = value
            '        Else
            '            _earliestStartDate = Heute
            '        End If

            '        If _Start > 0 Then
            '            _earliestStart = System.Math.Min(DateDiff(DateInterval.Month, _startDate, _earliestStartDate), 0)
            '        End If
            '    Else
            '        Throw New ArgumentException("unzulässiges frühestes Startdatum: " & value.ToString)
            '    End If

            'End If

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

            _earliestStartDate = value
            'If DateDiff(DateInterval.Month, StartofCalendar, value) + 1 > 0 And DateDiff(DateInterval.Month, _startDate, value) <= 0 Then
            '    If DateDiff(DateInterval.Month, Heute, value) > 0 Then
            '        _earliestStartDate = value
            '    Else
            '        _earliestStartDate = Heute
            '    End If

            '    If _Start > 0 Then
            '        _earliestStart = System.Math.Min(DateDiff(DateInterval.Month, _startDate, _earliestStartDate), 0)
            '    End If
            'Else
            '    Throw New ArgumentException("unzulässiges frühestes Startdatum: " & value.ToString)
            'End If



        End Set
    End Property

    Public Property latestStartDate As Date
        Get
            latestStartDate = _latestStartDate
        End Get
        Set(value As Date)
            Dim heute As Date = Now

            _latestStartDate = value
            'If _Status = ProjektStatus(1) Or _Status = ProjektStatus(2) Or _
            '                                 _Status = ProjektStatus(2) Then
            '    Throw New ArgumentException("der Startzeitpunkt kann nicht mehr verändert werden ... ")
            'Else

            '    If DateDiff(DateInterval.Month, StartofCalendar, value) + 1 > 0 And DateDiff(DateInterval.Month, _startDate, value) >= 0 Then
            '        If DateDiff(DateInterval.Month, heute, value) > 0 Then
            '            _latestStartDate = value
            '        Else
            '            _latestStartDate = heute
            '        End If
            '        If _Start > 0 Then
            '            _latestStart = System.Math.Max(DateDiff(DateInterval.Month, _startDate, _latestStartDate), 0)
            '        End If
            '    Else
            '        Throw New ArgumentException("unzulässiges spätestes Startdatum: " & value.ToString)
            '    End If
            'End If


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

            _latestStartDate = value

            'If DateDiff(DateInterval.Month, StartofCalendar, value) + 1 > 0 And DateDiff(DateInterval.Month, _startDate, value) >= 0 Then
            '    If DateDiff(DateInterval.Month, heute, value) > 0 Then
            '        _latestStartDate = value
            '    Else
            '        _latestStartDate = heute
            '    End If
            '    If _Start > 0 Then
            '        _latestStart = System.Math.Max(DateDiff(DateInterval.Month, _startDate, _latestStartDate), 0)
            '    End If
            'Else
            '    Throw New ArgumentException("unzulässiges spätestes Startdatum: " & value.ToString)
            'End If



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
            tmp = (Me.Risiko - weightStrategicFit * Me.StrategicFit) / 100
            If tmp < 0 Then
                tmp = 0
            End If
            risikoKostenfaktor = tmp
        End Get
    End Property
    ''' <summary>
    ''' kopiert die Attribute eines Projektes in newproject; unterscheidet dabei , ob es sich bei der Quelle um eine 
    ''' Vorlage handelt oder ein normales Projekt 
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <param name="isVorlage"></param>
    ''' <remarks></remarks>
    Public Sub copyAttrTo(ByRef newproject As clsProjekt, isVorlage As Boolean)

        With newproject
            .farbe = farbe
            .Schrift = Schrift
            .Schriftfarbe = Schriftfarbe
            .VorlagenName = VorlagenName
            .Risiko = Risiko
            .StrategicFit = StrategicFit
            .Erloes = Erloes



            If isVorlage Then
                .earliestStart = _earliestStart
                .latestStart = _latestStart
                .name = ""

            Else
                .StartOffset = _StartOffset
                .startDate = _startDate
                .earliestStartDate = _earliestStartDate
                .latestStartDate = _latestStartDate
                .earliestStart = _earliestStart
                .latestStart = _latestStart
                .leadPerson = ""
                '.ProjectMarge = imarge
            End If

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

                                Try
                                    If result.getBewertung(1).colorIndex = colorIndex Then
                                        resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                    End If
                                Catch ex1 As Exception
                                    ' hierher kommt er, wenn es ein Result, aber keine Bewertung gibt 
                                    If colorIndex = 0 Then
                                        resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                    End If
                                End Try



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
            Dim phasenStart As Integer, phasenEnde As Integer
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
                        phasenStart = .relStart - 1
                        phasenEnde = .relEnde - 1


                        For r = 1 To anzResults

                            result = .getResult(r)
                            monatsIndex = DateDiff(DateInterval.Month, Me.startDate, result.getDate)
                            ' Sicherstellen, daß Ergebnisse, die vor oder auch nach dem Projekt erreicht werden sollen, richtig behandelt werden 

                            If monatsIndex < 0 Then
                                monatsIndex = 0
                            ElseIf monatsIndex > Me.Dauer - 1 Then
                                monatsIndex = Me.Dauer - 1
                            End If


                            ResultValues(monatsIndex) = ResultValues(monatsIndex) & vbLf & result.name & _
                                                        " (" & result.getDate.ToShortDateString & ")"


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
                    newphase.AddResult(newresult)
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
                    newphase.AddResult(newresult)
                Next

            Catch ex As Exception
                ' in diesem Falle gibt es die komplette Phase in dem Projekt nicht mehr 
                ' dann muss auch nichts gemacht werden 
            End Try


        Next

    End Sub


    Public Overrides Sub CopyTo(ByRef newproject As clsProjekt)

        Dim newphase As clsPhase

        Call copyAttrTo(newproject, True)

        For Each hphase In MyBase.Liste
            newphase = New clsPhase(newproject)
            hphase.CopyTo(newphase)
            newproject.AddPhase(newphase)
        Next


    End Sub

    Public Overloads Sub CopyTo(ByRef newproject As clsProjekt, isVorlage As Boolean)

        Dim newphase As clsPhase

        Call copyAttrTo(newproject, isVorlage)

        For Each hphase In MyBase.Liste
            newphase = New clsPhase(newproject)
            hphase.CopyTo(newphase)
            newproject.AddPhase(newphase)
        Next


    End Sub


    '
    ' übergibt in Project Marge die berechnete Marge: Erloes - Kosten
    '
    Public ReadOnly Property ProjectMarge() As Double


        Get
            Dim gk As Double = Me.getGesamtKostenBedarf.Sum
            ' prüfen , ob die Marge konsistent ist mit Verhältnis Erlös und Kosten  ... 

            If gk > 0 Then
                ProjectMarge = (Me.Erloes - gk) / gk
            Else
                ProjectMarge = 0
            End If


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
            If value <= 0 Then
                _earliestStart = value
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

        'Set(value As Integer)

        '    Dim newDate As Date = StartofCalendar.AddMonths(value - 1)
        '    Dim Heute As Date = Now

        '    If _Start <> value Then

        '        If _Status = ProjektStatus(0) Then
        '            ' nur dann darf das Projekt verschoben werden  
        '            ' andernfalls muss nichts gemacht werden ...
        '            _Start = value
        '            '_tfSpalte = value
        '            Me.startDate = StartofCalendar.AddMonths(value - 1)
        '            Me.earliestStartDate = _startDate
        '            Me.latestStartDate = _startDate
        '        End If


        '        'Me.earliestStart = 0
        '        'Me.latestStart = 0



        '        'If _Start = 0 Then
        '        '    ' es handelt sich um die erstbesetzung ... 
        '        '    ' hier muss überprüft werden, ob die earliest/latest Settings so bleiben können: sie könnten ja bereits in der Vergangenheit liegen 
        '        '    _Start = value
        '        '    _startDate = StartofCalendar.AddMonths(value - 1)
        '        '    _tfSpalte = value

        '        '    'If DateDiff(DateInterval.Month, Heute, newDate) <= 0 Then
        '        '    '    _earliestStart = 0
        '        '    'ElseIf DateDiff(DateInterval.Month, newDate, Heute) > _earliestStart Then
        '        '    '    _earliestStart = DateDiff(DateInterval.Month, newDate, Heute)
        '        '    'End If

        '        'ElseIf _Status = ProjektStatus(1) Or _Status = ProjektStatus(2) Or _
        '        '                                 _Status = ProjektStatus(2) Then
        '        '    'Call MsgBox("der Startzeitpunkt kann nicht mehr verändert werden ... ")
        '        '    Throw New ApplicationException("der Startzeitpunkt kann nicht mehr verändert werden ... ")

        '        'ElseIf value < _Start + _earliestStart Then
        '        '    'Call MsgBox("der neue Startzeitpunkt liegt vor dem bisher zugelassenen frühestmöglichen Startzeitpunkt ...")
        '        '    Throw New ApplicationException("der neue Startzeitpunkt liegt vor dem bisher zugelassenen frühestmöglichen Startzeitpunkt ...")

        '        'ElseIf value > _Start + _latestStart Then
        '        '    'Call MsgBox("der neue Startzeitpunkt liegt nach dem bisher zugelassenen spätestmöglichen Startzeitpunkt ...")
        '        '    Throw New ApplicationException("der neue Startzeitpunkt liegt nach dem bisher zugelassenen spätestmöglichen Startzeitpunkt ...")
        '        'Else

        '        '    If DateDiff(DateInterval.Month, Heute, newDate) < 0 Then
        '        '        'Call MsgBox("der neue Startzeitpunkt liegt in der Vergangenheit ...")
        '        '        Throw New ApplicationException("der neue Startzeitpunkt liegt in der Vergangenheit ...")
        '        '    Else

        '        '        _Start = value
        '        '        _tfSpalte = value
        '        '        Me.startDate = StartofCalendar.AddMonths(value - 1)
        '        '        Me.earliestStart = System.Math.Min(DateDiff(DateInterval.Month, _startDate, _earliestStartDate), 0)
        '        '        Me.latestStart = System.Math.Max(DateDiff(DateInterval.Month, _startDate, _latestStartDate), 0)

        '        '    End If


        '        'End If

        '    End If
        'End Set
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
                _earliestStart = 0
                _latestStart = 0
                _earliestStartDate = _startDate
                _latestStartDate = _startDate
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
                    top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + 0.1 * boxHeight
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


    Public Sub calculateResultCoord(ByVal resultDate As Date, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        'Dim endDatum As Date = StartofCalendar.AddMonths(Me.Start - 1 + Dauer).AddDays(-1)
        Dim diffMonths As Integer = DateDiff(DateInterval.Month, StartofCalendar, resultDate)
        Dim dayOfResult As Integer = resultDate.Day

        'Dim tagebisResult As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start - 1), resultDate)
        'Dim ratio As Double = tagebisResult / anzahlTage

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.Dauer > 0 Then
            top = topOfMagicBoard + (Me.tfZeile - 1.0) * boxHeight - boxWidth / 2
            left = diffMonths * boxWidth + dayOfResult * (boxWidth / 30) - 0.5 * boxWidth

            'width = 0.66 * boxWidth
            'height = 0.66 * boxWidth
            width = boxWidth * 1.1
            height = boxWidth * 1.1
        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub

    Public Sub calculateStatusCoord(ByVal resultDate As Date, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        'Dim endDatum As Date = StartofCalendar.AddMonths(Me.Start - 1 + Dauer).AddDays(-1)
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
        _timeStamp = Date.Now

    End Sub

    Public Sub New(ByVal projektStart As Integer, ByVal earliestValue As Integer, ByVal latestValue As Integer)

        AllPhases = New List(Of clsPhase)
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
