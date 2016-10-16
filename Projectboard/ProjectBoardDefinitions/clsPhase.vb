Imports Microsoft.Office.Interop.Excel
Public Class clsPhase

    ' earliestStart und latestStart sind absolute Werte im "koordinaten-System" des Projektes
    ' von daher ist es anders gelöst als in clsProjekt, wo earlieststart und latestStart relative Angaben sind 

    Private _nameID As String
    Private _parentProject As clsProjekt
    Private _vorlagenParent As clsProjektvorlage

    Private _shortName As String
    Private _originalName As String
    Private _appearance As String
    Private _color As Integer

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



    ''' <summary>
    ''' liest/schreibt das Feld für vrantwortlich
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
    Public Property appearance As String
        Get
            appearance = _appearance
        End Get
        Set(value As String)
            If appearanceDefinitions.ContainsKey(value) Then
                _appearance = value
            Else
                _appearance = awinSettings.defaultPhaseClass
            End If
        End Set
    End Property

    ''' <summary>
    ''' gibt das Shape für die Phase zurück
    ''' falls es keine explizite Definition gibt: die Form der ersten Phase in der AppearnceDefinitions-Liste 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShape As Microsoft.Office.Interop.Excel.Shape
        Get

            Dim tmpClass As String
            Dim found As Boolean = True

            If PhaseDefinitions.Contains(Me.name) Then
                tmpClass = PhaseDefinitions.getPhaseDef(Me.name).darstellungsKlasse

            ElseIf missingMilestoneDefinitions.Contains(Me.name) Then
                tmpClass = missingPhaseDefinitions.getPhaseDef(Me.name).darstellungsKlasse

            Else
                tmpClass = _appearance
                found = False
            End If

            getShape = appearanceDefinitions.Item(tmpClass).form

            If Not found Then
                getShape.Fill.ForeColor.RGB = _color
            End If

        End Get
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
                ampelStatus = Me.getBewertung(1).colorIndex
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
                Me.getBewertung(1).colorIndex = value
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
                ampelErlaeuterung = Me.getBewertung(1).description
            Else
                ampelErlaeuterung = ""
            End If
        End Get
        Set(value As String)
            If IsNothing(value) Then
                value = ""
            End If

            If Me.bewertungsCount >= 1 Then
                Me.getBewertung(1).description = value
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

            projektStartdate = Me.parentProject.startDate
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
        Dim faktor As Double
        Dim dimension As Integer


        If dauer < 0 Then
            Throw New ArgumentException("Dauer kann nicht negativ sein")

        ElseIf startOffset < 0 Then
            Throw New ArgumentException("Phase kann nicht vor Projektstart beginnen")

        End If


        Try

            projektStartdate = Me.parentProject.startDate
            projektstartColumn = Me.parentProject.Start

            If dauer = 0 And _relEnde > 0 Then

                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = CInt(DateDiff(DateInterval.Day, projektStartdate, projektStartdate.AddMonths(_relStart - 1)))
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            ElseIf dauer = 0 And _relEnde = 0 Then

                Throw New ArgumentException("Phase kann nicht Dauer = 0 haben ")

            Else
                '  

                _startOffsetinDays = startOffset
                _dauerInDays = dauer

                Dim oldlaenge As Integer = _relEnde - _relStart + 1


                Dim phaseStartdate As Date = Me.getStartDate
                Dim phaseEndDate As Date = Me.getEndDate


                _relStart = getColumnOfDate(phaseStartdate) - projektstartColumn + 1
                _relEnde = getColumnOfDate(phaseEndDate) - projektstartColumn + 1


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
                abbrev = _shortName
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

                If PhaseDefinitions.Contains(phName) Or missingPhaseDefinitions.Contains(phName) Then
                    farbe = Me.getShape.Fill.ForeColor.RGB
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

                If Me.parentProject Is Nothing Then
                    isVorlage = True
                Else
                    isVorlage = False
                End If
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

                If Me.parentProject Is Nothing Then
                    isVorlage = True
                Else
                    isVorlage = False
                End If
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


            If startpunkt < 0 Then
                Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
            End If

            Dim phasenStart As Integer = startpunkt + Me.startOffsetinDays
            Dim phasenDauer As Integer = Me.dauerInDays



            If tfzeile > 1 And phasenStart >= 1 And phasenDauer > 0 Then


                top = topOfMagicBoard + (tfzeile - 1) * boxHeight + 0.5 * (0.8 - 0.23) * boxHeight
                left = (phasenStart / 365) * boxWidth * 12
                width = ((phasenDauer) / 365) * boxWidth * 12
                height = 0.23 * boxHeight

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
    ''' gibt die Rollen Instanz der Phase zurück, die den Namen roleName hat 
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRole(ByVal roleName As String) As clsRolle

        Get
            Dim returnValue As clsRolle = Nothing
            Dim ix As Integer = 0
            Dim found As Boolean = False

            While Not found And ix <= _allRoles.Count - 1
                If _allRoles.Item(ix).name = roleName Then
                    found = True
                    returnValue = _allRoles.Item(ix)
                Else
                    ix = ix + 1
                End If
            End While

            getRole = returnValue

        End Get

    End Property


    ''' <summary>
    ''' addRole fügt die Rollen Instanz hinzu, wenn sie nicht schon existiert
    ''' summiert die Werte zu der shon existierenden ...
    ''' </summary>
    ''' <param name="role"></param>
    ''' <remarks></remarks>
    Public Sub addRole(ByVal role As clsRolle)

        'sollte nach dem 8.7.16 aktiviert werden 
        'ebenso für addCost, mehrere Rollen/Kosten des gleichen NAmens sollen aufsummiert werden 
        Dim roleName As String = role.name
        Dim returnValue As clsRolle = Nothing
        Dim ix As Integer = 0
        Dim found As Boolean = False
        Dim oldXWerte() As Double
        Dim newXwerte() As Double

        While Not found And ix <= _allRoles.Count - 1
            If _allRoles.Item(ix).name = roleName Then
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
        Me.parentProject.rcLists.addRP(role.RollenTyp, Me.nameID)

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
    Public Sub removeRoleByName(ByVal roleName As String)

        Dim toDoList As New List(Of clsRolle)

        For i As Integer = 1 To _allRoles.Count
            Dim tmpRole As clsRolle = _allRoles.Item(i - 1)
            If tmpRole.name = roleName Then
                toDoList.Add(tmpRole)
            End If
        Next

        For Each tmpRole As clsRolle In toDoList
            _allRoles.Remove(tmpRole)
            ' Änderung tk 20.09.16
            ' jetzt müssen die sortierten Listen im Projekt entsprechend aktualisiert werden 
            Me.parentProject.rcLists.removeRP(tmpRole.RollenTyp, Me.nameID)
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



    Public Sub copyTo(ByRef newphase As clsPhase)
        Dim r As Integer, k As Integer
        Dim newrole As clsRolle
        Dim newcost As clsKostenart
        Dim newresult As clsMeilenstein
        ' Dimension ist die Länge des Arrays , der kopiert werden soll; 
        ' mit der eingeführten Unschärfe ist nicht mehr gewährleistet, 
        ' daß relende-relstart die tatsächliche Dimension des Arrays wiedergibt 
        Dim dimension As Integer

        With newphase

            .earliestStart = Me._earliestStart
            .latestStart = Me._latestStart
            .offset = Me._offset
            .nameID = _nameID

            ' sonstigen Elemente übernehmen 
            .shortName = Me._shortName
            .originalName = Me._originalName
            .appearance = Me._appearance
            .farbe = Me._color
            .verantwortlich = Me._verantwortlich

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


            ' Änderung 16.1.2014: zuerst die Rollen und Kosten übertragen, dann die relStart und RelEnde, dann die Results
            ' die evtl. enstehende Inkonsistenz zwischen Längen der Arrays der Rollen/Kostenarten und dem neuen relende/relstart wird in Kauf genommen 
            ' und nur korrigiert , wenn explizit gewünscht (Parameter awinsettings.autoCorrectBedarfe = true 

            .changeStartandDauer(Me._startOffsetinDays, Me._dauerInDays)

            For r = 1 To Me._allMilestones.Count
                newresult = New clsMeilenstein(parent:=newphase)
                Me.getMilestone(r).copyTo(newresult)

                Try
                    .addMilestone(newresult)
                Catch ex As Exception

                End Try

            Next


            ' jetzt noch die evtl vorhandenen Bewertungen kopieren 
            For b As Integer = 1 To Me._bewertungen.Count
                Dim newb As New clsBewertung
                Me.getBewertung(b).copyto(newb)
                Try
                    .addBewertung(newb)
                Catch ex As Exception

                End Try

            Next

        End With

    End Sub
    Public Sub korrCopyTo(ByRef newphase As clsPhase, ByVal corrFactor As Double, Optional newPhaseNameID As String = "")
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


            .changeStartandDauer(CInt(Me._startOffsetinDays * corrFactor), CInt(Me._dauerInDays * corrFactor))

            For r = 1 To Me.countRoles
                Try

                    oldrole = Me.getRole(r)
                    dimension = newphase.relEnde - newphase.relStart
                    newrole = New clsRolle(dimension)
                    ReDim newXwerte(dimension)
                    hname = oldrole.name

                    Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldrole.Xwerte, corrFactor, newXwerte)

                    With newrole
                        .RollenTyp = oldrole.RollenTyp
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

                    Call berechneBedarfe(newphase.getStartDate.Date, newphase.getEndDate.Date, oldcost.Xwerte, corrFactor, newXwerte)

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

    Public Property Role(ByVal index As Integer) As clsRolle
        Get
            Role = _allRoles.Item(index - 1)
        End Get

        Set(value As clsRolle)
            _allRoles.Item(index - 1) = value
        End Set

    End Property

    Public Property Cost(ByVal index As Integer) As clsKostenart
        Get
            Cost = _allCosts.Item(index - 1)
        End Get

        Set(value As clsKostenart)
            _allCosts.Item(index - 1) = value
        End Set

    End Property

    ''' <summary>
    ''' liefert die Rolle an Index-Stelle i; i darf Werte zwischen 1 und AnzahlRollen annehmen
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRole(ByVal index As Integer) As clsRolle

        Get
            getRole = _allRoles.Item(index - 1)
        End Get

    End Property

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
            getCost = _allCosts.Item(index - 1)
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

        _bewertungen = New SortedList(Of String, clsBewertung)
        _allRoles = New List(Of clsRolle)
        _allCosts = New List(Of clsKostenart)
        _allMilestones = New List(Of clsMeilenstein)

        _shortName = ""
        _originalName = ""
        _appearance = awinSettings.defaultPhaseClass

        Try
            _color = XlRgbColor.rgbDarkGrey
            If appearanceDefinitions.ContainsKey(_appearance) Then
                If Not IsNothing(appearanceDefinitions.Item(_appearance).form) Then
                    _color = appearanceDefinitions.Item(_appearance).form.Fill.ForeColor.RGB
                End If
            End If

        Catch ex As Exception

        End Try

        _verantwortlich = ""

        _offset = 0
        _earliestStart = -999
        _latestStart = -999
        




    End Sub

    Public Sub New(ByRef parent As clsProjektvorlage, ByVal isVorlage As Boolean)
        ' Variable isVorlage dient lediglich dazu, eine weitere Signatur für einen Konstruktor zu bekommen 
        ' dieser Konstruktor wird für parent = Vorlage benutzt 

        _nameID = ""
        _parentProject = Nothing
        _vorlagenParent = parent


        _bewertungen = New SortedList(Of String, clsBewertung)
        _allRoles = New List(Of clsRolle)
        _allCosts = New List(Of clsKostenart)
        _allMilestones = New List(Of clsMeilenstein)

        _shortName = ""
        _originalName = ""
        _appearance = awinSettings.defaultPhaseClass

        Try
            _color = XlRgbColor.rgbDarkGrey
            If appearanceDefinitions.ContainsKey(_appearance) Then
                If Not IsNothing(appearanceDefinitions.Item(_appearance).form) Then
                    _color = appearanceDefinitions.Item(_appearance).form.Fill.ForeColor.RGB
                End If
            End If

        Catch ex As Exception

        End Try

        _verantwortlich = ""
        
        _offset = 0
        _earliestStart = -999
        _latestStart = -999
        

    End Sub

    ''' <summary>
    ''' synchronisiert bzw. berechnet die Xwerte der Rollen und Kosten
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub calcNewXwerte(ByVal dimension As Integer, ByVal faktor As Double)
        Dim newXwerte() As Double
        Dim oldXwerte() As Double

        Dim r As Integer, k As Integer

        For r = 1 To Me.countRoles
            oldXwerte = Me.getRole(r).Xwerte
            ReDim newXwerte(dimension)
            Call berechneBedarfe(Me.getStartDate.Date, Me.getEndDate.Date, oldXwerte, faktor, newXwerte)
            Me.getRole(r).Xwerte = newXwerte
        Next

        For k = 1 To Me.countCosts
            oldXwerte = Me.getCost(k).Xwerte
            ReDim newXwerte(dimension)
            Call berechneBedarfe(Me.getStartDate.Date, Me.getEndDate.Date, oldXwerte, faktor, newXwerte)
            Me.getCost(k).Xwerte = newXwerte
        Next


    End Sub


    ''' <summary>
    ''' berechnet die Bedarfe (Rollen,Kosten) der Phase gemäß Startdate und endedate, und corrFakt neu
    ''' </summary>
    ''' <param name="startdate"></param>
    ''' <param name="endedate"></param>
    ''' <param name="oldXwerte"></param>
    ''' <param name="corrFakt"></param>
    ''' <param name="newValues"></param>
    ''' <remarks></remarks>
    Public Sub berechneBedarfe(ByVal startdate As Date, ByVal endedate As Date, ByVal oldXwerte() As Double, _
                               ByVal corrFakt As Double, ByRef newValues() As Double)
        Dim k As Integer
        Dim newXwerte() As Double
        Dim gesBedarf As Double
        Dim Rest As Integer
        Dim hDatum As Date
        Dim anzDaysthisMonth As Double

        Try
            ReDim newXwerte(newValues.Length - 1)

            gesBedarf = oldXwerte.Sum
            If awinSettings.propAnpassRess Then
                ' Gesamter Bedarf dieser Rolle/Kosten wird gemäß streckung bzw. stauchung des Projekts korrigiert
                gesBedarf = System.Math.Round(gesBedarf * corrFakt)
            End If

            If newValues.Length = oldXwerte.Length Then

                'Bedarfe-Verteilung bleibt wie gehabt, aber die corrfakt ist hier unberücksichtigt ..? 

                If Not awinSettings.propAnpassRess Then
                    newXwerte = oldXwerte
                Else
                    For i = 0 To newValues.Length - 1
                        newXwerte(i) = System.Math.Round(oldXwerte(i) * corrFakt)
                    Next

                    ' jetzt ggf die Reste verteilen 
                    Rest = CInt(System.Math.Round(oldXwerte.Sum * corrFakt - newXwerte.Sum))

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

            End If

            newValues = newXwerte

        Catch ex As Exception

            Call MsgBox("Fehler in berechneBedarfe: " & vbLf & ex.Message)

        End Try




    End Sub

End Class
