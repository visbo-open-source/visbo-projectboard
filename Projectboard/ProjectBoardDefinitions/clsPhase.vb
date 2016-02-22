Public Class clsPhase

    ' earliestStart und latestStart sind absolute Werte im "koordinaten-System" des Projektes
    ' von daher ist es anders gelöst als in clsProjekt, wo earlieststart und latestStart relative Angaben sind 

    Private AllMilestones As List(Of clsMeilenstein)
    Private AllRoles As List(Of clsRolle)
    Private AllCosts As List(Of clsKostenart)
    Private _Offset As Integer
    Private _earliestStart As Integer
    Private _latestStart As Integer
    Private _minDauer As Integer
    Private _maxDauer As Integer
    Private _relStart As Integer
    Private _relEnde As Integer
    Private _name As String
    Private _startOffsetinDays As Integer
    Private _dauerInDays As Integer
    Private _Parent As clsProjekt
    Private _vorlagenParent As clsProjektvorlage

    ' Erweiterung tk 18.2.16
    ' das wird verwendet . um eine Farbe Meilensteins, der nicht zur Liste der bekannten gehört 
    ' aufzunehmen 
    Private _alternativeColor As Long


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

            projektStartdate = Me.Parent.startDate
            projektstartColumn = Me.Parent.Start

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
                    If Me.nameID <> Me.Parent.getPhase(1).nameID Then
                        ' wenn es nicht die erste Phase ist, die gerade behandelt wird, dann soll die erste Phase auf Konsistenz geprüft werden 
                        Me.Parent.keepPhase1consistent(Me.startOffsetinDays + Me.dauerInDays)
                    End If
                Catch ex As Exception

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

            projektStartdate = Me.Parent.startDate
            projektstartColumn = Me.Parent.Start

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


    Public Property Offset As Integer
        Get
            Offset = _Offset
        End Get
        Set(value As Integer)
            If _earliestStart = -999 Or _latestStart = -999 Then
                _Offset = value
            Else
                If value >= _earliestStart - _relStart And value <= _latestStart - _relStart Then
                    _Offset = value
                Else
                    Throw New ApplicationException("Wert für Offset liegt ausserhalb der zugelassenen Grenzen")
                End If
            End If

        End Set
    End Property

    ''' <summary>
    ''' gibt den Original Namen einer Phase zurück 
    ''' wenn der leer ist, dann wird der Phasen Name zurück gegeben 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property originalName As String
        Get

            Dim tmpNode As clsHierarchyNode
            Dim beschriftung As String = Me.name
            tmpNode = _Parent.hierarchy.nodeItem(Me.nameID)

            If Not IsNothing(tmpNode) Then
                beschriftung = tmpNode.origName
                If beschriftung = "" Then
                    beschriftung = Me.name
                End If
            Else
                beschriftung = Me.name
            End If

            originalName = beschriftung

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
            getStartDate = Me.Parent.startDate.AddDays(_startOffsetinDays)
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
                getEndDate = Me.Parent.startDate.AddDays(_startOffsetinDays + _dauerInDays - 1)
            Else
                'Throw New Exception("Dauer muss mindestens 1 Tag sein ...")
                getEndDate = Me.Parent.startDate.AddDays(_startOffsetinDays)
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die Farbe einer Phase zurück; das ist die Farbe der Darstellungsklasse, wenn die Phase zur Liste der
    ''' bekannten Elemente gehört, sonst die AlternativeFare, die ggf beim auslesen z.b. aus MS Project ermittelt wird
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property farbe As Object
        Get
            Try
                Dim itemName As String = elemNameOfElemID(_name)
                If _name = rootPhaseName Then
                    farbe = Me.Parent.farbe             ' Farbe der Projektes, da Projekt der Parent der RootPhase ist
                Else
                    Dim tmpPhaseDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(elemNameOfElemID(_name))
                    If IsNothing(tmpPhaseDef) Then
                        farbe = _alternativeColor
                    Else
                        farbe = tmpPhaseDef.farbe
                    End If

                End If

            Catch ex As Exception
                ' in diesem Fall wird ein Standard Farbe genommen 
                farbe = awinSettings.AmpelNichtBewertet
            End Try

        End Get
    End Property


    ''' <summary>
    ''' setzt die Farbe eines Meilensteins; macht  dann Sinn, wenn der Meilenstein nicht zur 
    ''' Liste der bekannten Meilensteine gehört 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property setFarbe As Long
        Set(value As Long)

            If value >= RGB(0, 0, 0) And value <= RGB(255, 255, 255) Then
                _alternativeColor = value
            Else
                ' unverändert lassen - wird ja auch im New initial gesetzt 
            End If

        End Set
    End Property


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

    Public Property minDauer As Integer
        Get
            minDauer = _minDauer
        End Get
        Set(value As Integer)
            If value >= 1 Then
                If _maxDauer <> -999 Then
                    If value <= _maxDauer Then
                        _minDauer = value
                    Else
                        Throw New ApplicationException("Mindest-Dauer kann nicht größer als Max Dauer sein")
                    End If
                Else
                    _minDauer = value
                End If
            Else
                Throw New ApplicationException("Mindest-Dauer kann nicht negativ oder Null sein")
            End If

        End Set
    End Property

    Public Property maxDauer As Integer
        Get
            maxDauer = _maxDauer
        End Get
        Set(value As Integer)
            If value >= 1 Then
                If _minDauer <> -999 Then
                    If value >= _minDauer Then
                        _maxDauer = value
                    Else
                        Throw New ApplicationException("Maximal-Dauer kann nicht kleiner als Min Dauer sein")
                    End If
                Else
                    _maxDauer = value
                End If
            Else
                Throw New ApplicationException("Maximal-Dauer kann nicht negativ oder Null sein")
            End If

        End Set
    End Property


    Public ReadOnly Property relStart As Integer
        Get

            Dim isVorlage As Boolean
            Dim tmpValue As Integer
            'Dim checkValue As Integer = _relStart + _Offset

            Try

                If Me.Parent Is Nothing Then
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
                tmpValue = getColumnOfDate(Me.Parent.startDate.AddDays(Me.startOffsetinDays)) - Me.Parent.Start + 1
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

                If Me.Parent Is Nothing Then
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
                tmpValue = getColumnOfDate(Me.Parent.startDate.AddDays(Me.startOffsetinDays + Me.dauerInDays - 1)) - Me.Parent.Start + 1
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
            nameID = _name
        End Get
        Set(value As String)
            Dim tmpstr() As String
            tmpstr = value.Split(New Char() {CChar("§")}, 3)
            If Len(value) > 0 Then
                If value.StartsWith("0§") And tmpstr.Length >= 2 Then
                    _name = value
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
            name = elemNameOfElemID(_name)
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

            Dim projektStartdate As Date = Me.Parent.startDate
            Dim tfzeile As Integer = Me.Parent.tfZeile
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



    'End Sub
    Public Sub addRole(ByVal role As clsRolle)

        If Not AllRoles.Contains(role) Then
            AllRoles.Add(role)
        Else
            'Call logfileSchreiben("Fehler: Rolle '" & role.name & "' ist bereits in der Phase '" & Me.name & "' enthalten", "", anzFehler)
        End If


    End Sub

    ''' <summary>
    ''' es wird überprüft, ob der Meilenstein-Name schon existiert 
    ''' wenn er bereits existiert, wird eine ArgumentException geworfen  
    ''' </summary>
    ''' <param name="milestone"></param>
    ''' <remarks></remarks>
    Public Sub addMilestone(ByVal milestone As clsMeilenstein,
                            Optional ByVal origName As String = "")


        Dim anzElements As Integer = AllMilestones.Count - 1
        Dim ix As Integer = 0
        Dim found As Boolean = False

        Dim elemName As String = elemNameOfElemID(milestone.nameID)

        Do While ix <= anzElements And Not found
            If AllMilestones.Item(ix).nameID = milestone.nameID Then
                found = True
            Else
                ix = ix + 1
            End If
        Loop

        If found Then
            Throw New ArgumentException("Meilenstein existiert bereits in dieser Phase!" & milestone.nameID)
        Else
            AllMilestones.Add(milestone)
        End If

        ' jetzt muss der Meilenstein in die Projekt-Hierarchie aufgenommen werden , 
        ' aber nur, wenn die Phase bereits in der Projekt-Hierarchie vorhanden ist ... 
        Dim elemID As String = milestone.nameID
        Dim currentElementNode As New clsHierarchyNode
        Dim hproj As New clsProjekt, vproj As New clsProjektvorlage
        Dim parentIsVorlage As Boolean
        Dim milestoneIndex As Integer = AllMilestones.Count
        Dim phaseID As String = Me.nameID
        Dim ok As Boolean = False

        If Not istElemID(elemID) Then
            elemID = vproj.hierarchy.findUniqueElemKey(elemName, True)
        End If

        If IsNothing(Me.Parent) Then
            parentIsVorlage = True
            vproj = Me.VorlagenParent
            If vproj.hierarchy.containsKey(phaseID) Then
                ' Phase ist bereits in der Projekt-Hierarchie eingetragen
                ok = True
            End If
        Else
            parentIsVorlage = False
            hproj = Me.Parent
            If hproj.hierarchy.containsKey(phaseID) Then
                ' Phase ist bereits in der Projekt-Hierarchie eingetragen
                ok = True
            End If
        End If

        If ok Then

            With currentElementNode

                .elemName = elemName

                If origName = "" Then
                    .origName = .elemName
                Else
                    .origName = origName
                End If

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

        If index >= 0 And index <= AllMilestones.Count - 1 Then
            If checkID <> "" Then
                If AllMilestones.ElementAt(index).nameID = checkID Then
                    ok = True
                Else
                    ok = False
                End If
            End If
        Else
            ok = False
        End If
        

        If ok Then
            AllMilestones.RemoveAt(index)
        End If

    End Sub

    Public ReadOnly Property rollenListe() As List(Of clsRolle)

        Get
            rollenListe = AllRoles
        End Get

    End Property

    Public ReadOnly Property meilensteinListe() As List(Of clsMeilenstein)

        Get
            meilensteinListe = AllMilestones
        End Get

    End Property

    Public ReadOnly Property kostenListe() As List(Of clsKostenart)

        Get
            kostenListe = AllCosts
        End Get

    End Property


    Public ReadOnly Property countRoles() As Integer

        Get
            countRoles = AllRoles.Count
        End Get

    End Property

    Public ReadOnly Property countMilestones() As Integer

        Get
            countMilestones = AllMilestones.Count
        End Get

    End Property



    Public Sub CopyTo(ByRef newphase As clsPhase)
        Dim r As Integer, k As Integer
        Dim newrole As clsRolle
        Dim newcost As clsKostenart
        Dim newresult As clsMeilenstein
        ' Dimension ist die Länge des Arrays , der kopiert werden soll; 
        ' mit der eingeführten Unschärfe ist nicht mehr gewährleistet, 
        ' daß relende-relstart die tatsächliche Dimension des Arrays wiedergibt 
        Dim dimension As Integer

        With newphase
            .minDauer = Me._minDauer
            .maxDauer = Me._maxDauer
            .earliestStart = Me._earliestStart
            .latestStart = Me._latestStart
            .Offset = Me._Offset



            .nameID = _name

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

            For r = 1 To Me.AllMilestones.Count
                newresult = New clsMeilenstein(parent:=newphase)
                Me.getMilestone(r).CopyTo(newresult)

                Try
                    .addMilestone(newresult)
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
            .minDauer = Me._minDauer
            .maxDauer = Me._maxDauer
            .earliestStart = Me._earliestStart
            .latestStart = Me._latestStart
            .Offset = Me._Offset

            If newPhaseNameID = "" Then
                .nameID = _name
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

            For r = 1 To Me.AllMilestones.Count
                newresult = New clsMeilenstein(parent:=newphase)
                If newPhaseNameID = "" Then
                    Me.getMilestone(r).CopyTo(newresult)
                Else
                    Dim newMSNameID As String = newphase.Parent.hierarchy.findUniqueElemKey(Me.getMilestone(r).name, True)
                    Me.getMilestone(r).CopyTo(newresult, newMSNameID)
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
        For r = 1 To Me.AllMilestones.Count

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
            Role = AllRoles.Item(index - 1)
        End Get

        Set(value As clsRolle)
            AllRoles.Item(index - 1) = value
        End Set

    End Property

    Public Property Cost(ByVal index As Integer) As clsKostenart
        Get
            Cost = AllCosts.Item(index - 1)
        End Get

        Set(value As clsKostenart)
            AllCosts.Item(index - 1) = value
        End Set

    End Property

    Public ReadOnly Property getRole(ByVal index As Integer) As clsRolle

        Get
            getRole = AllRoles.Item(index - 1)
        End Get

    End Property

    Public ReadOnly Property getMilestone(ByVal index As Integer) As clsMeilenstein

        Get
            If index < 1 Or index > AllMilestones.Count Then
                getMilestone = Nothing
            Else
                getMilestone = AllMilestones.Item(index - 1)
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

                hryNode = Me.Parent.hierarchy.nodeItem(key)
                If Not IsNothing(hryNode) Then

                    ' prüfen, ob der Meilenstein überhaupt zu dieser Phase gehört 
                    If hryNode.parentNodeKey = Me.nameID Then
                        index = hryNode.indexOfElem
                        tmpMilestone = AllMilestones.Item(index - 1)
                    End If

                End If


            Else

                Dim r As Integer = 1
                While r <= Me.countMilestones And Not found

                    If elemNameOfElemID(AllMilestones.Item(r - 1).nameID) = key Then
                        anzahl = anzahl + 1
                        If anzahl >= lfdNr Then
                            found = True
                            tmpMilestone = AllMilestones.Item(r - 1)
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

    Public Sub AddCost(ByVal cost As clsKostenart)

        If Not AllCosts.Contains(cost) Then
            AllCosts.Add(cost)
        Else
            Throw New Exception("Fehler: Kostenart '" & cost.name & "' ist bereits in der Phase '" & Me.name & "' enthalten")
        End If

    End Sub


    Public ReadOnly Property countCosts() As Integer

        Get
            countCosts = AllCosts.Count
        End Get

    End Property


    Public ReadOnly Property getCost(ByVal index As Integer) As clsKostenart

        Get
            getCost = AllCosts.Item(index - 1)
        End Get

    End Property

    Public ReadOnly Property Parent() As clsProjekt
        Get
            Parent = _Parent
        End Get
    End Property

    Public ReadOnly Property VorlagenParent() As clsProjektvorlage
        Get
            VorlagenParent = _vorlagenParent
        End Get
    End Property

    Public Sub New(ByRef parent As clsProjekt)

        AllRoles = New List(Of clsRolle)
        AllCosts = New List(Of clsKostenart)
        AllMilestones = New List(Of clsMeilenstein)
        _minDauer = 1
        _maxDauer = 60
        _Offset = 0
        _earliestStart = -999
        _latestStart = -999
        _Parent = parent
        _vorlagenParent = Nothing

        _alternativeColor = awinSettings.AmpelNichtBewertet


    End Sub

    Public Sub New(ByRef parent As clsProjektvorlage, ByVal isVorlage As Boolean)
        ' Variable isVorlage dient lediglich dazu, eine weitere Signatur für einen Konstruktor zu bekommen 
        ' dieser Konstruktor wird für parent = Vorlage benutzt 

        Dim defaultName As String = "Phasen Default"
        AllRoles = New List(Of clsRolle)
        AllCosts = New List(Of clsKostenart)
        AllMilestones = New List(Of clsMeilenstein)
        _minDauer = 1
        _maxDauer = 60
        _Offset = 0
        _earliestStart = -999
        _latestStart = -999
        _Parent = Nothing
        _vorlagenParent = parent

        _alternativeColor = awinSettings.AmpelNichtBewertet
        



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
    Public Sub berechneBedarfe(ByVal startdate As Date, ByVal endedate As Date, ByVal oldXwerte() As Double, ByVal corrFakt As Double, ByRef newValues() As Double)
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


                For k = 0 To newXwerte.Length - 1

                    If k = 0 Then
                        ' damit ist 00:00 des Startdates gemeint 
                        hDatum = startdate
                        anzDaysthisMonth = DateDiff("d", hDatum, DateSerial(hDatum.Year, hDatum.Month + 1, hDatum.Day))
                        anzDaysthisMonth = anzDaysthisMonth - DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum) - 1

                    ElseIf k = newXwerte.Length - 1 Then
                        ' damit hDatum das End-Datum um 23.00 Uhr
                        hDatum = endedate.AddHours(23)
                        anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month, 1), hDatum)

                    Else
                        hDatum = startdate
                        anzDaysthisMonth = DateDiff("d", DateSerial(hDatum.Year, hDatum.Month + k, hDatum.Day), DateSerial(hDatum.Year, hDatum.Month + k + 1, hDatum.Day))
                    End If

                    newXwerte(k) = System.Math.Round(anzDaysthisMonth / (Me.dauerInDays * corrFakt) * gesBedarf)

                Next k

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
