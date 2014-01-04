Public Class clsPhase

    ' earliestStart und latestStart sind absolute Werte im "koordinaten-System" des Projektes
    ' von daher ist es anders gelöst als in clsProjekt, wo earlieststart und latestStart relative Angaben sind 

    Private AllResults As List(Of clsResult)
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


    Public Sub changeStartandDauer(ByVal startOffset As Integer, ByVal dauer As Integer)

        Dim projektStartdate As Date
        Dim projektstartColumn As Integer

        Dim phaseStartdate As Date
        Dim phaseEndDate As Date


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
                _startOffsetinDays = DateDiff(DateInterval.Day, projektStartdate, projektStartdate.AddMonths(_relStart - 1))
                _dauerInDays = DateDiff(DateInterval.Day, projektStartdate.AddMonths(_relStart - 1), _
                                        projektStartdate.AddMonths(_relEnde).AddDays(-1)) + 1


            ElseIf dauer = 0 And _relEnde = 0 Then

                Throw New ArgumentException("Phase kann nicht Dauer = 0 haben ")

            Else
                '  

                _startOffsetinDays = startOffset
                _dauerInDays = dauer

                Dim oldlaenge As Integer = _relEnde - _relStart + 1
                Dim newlaenge As Integer

                '_relStart = DateDiff(DateInterval.Month, projektStartdate, projektStartdate .AddDays(startOffset)) + 1
                '_relEnde = DateDiff(DateInterval.Month, projektStartdate, projektStartdate.AddDays(startOffset + _dauerInDays)) + 1

                phaseStartdate = projektStartdate.AddDays(startOffset)
                phaseEndDate = projektStartdate.AddDays(startOffset + _dauerInDays - 1)

                _relStart = DateDiff(DateInterval.Month, StartofCalendar, phaseStartdate) + 1 - projektstartColumn + 1
                _relEnde = DateDiff(DateInterval.Month, StartofCalendar, phaseEndDate) + 1 - projektstartColumn + 1


                newlaenge = _relEnde - _relStart + 1

                If newlaenge <> oldlaenge Then
                    Dim newvalues() As Double
                    Dim oldvalues() As Double


                    Try

                        For r = 1 To Me.CountRoles
                            oldvalues = Me.getRole(r).Xwerte
                            newvalues = adjustArrayLength(newlaenge, oldvalues, False)
                            ' wahrscheinlich muss dafür eine .XwerteReDim Property gemacht werden
                            ' die den Redim macht ... 
                            Me.getRole(r).Xwerte = newvalues
                        Next


                        For k = 1 To Me.CountCosts
                            oldvalues = Me.getCost(k).Xwerte
                            newvalues = adjustArrayLength(newlaenge, oldvalues, False)
                            ' wahrscheionlich muss dafür eine .XwerteReDim Property gemacht werden
                            ' die den Redim macht ... 
                            Me.getCost(k).Xwerte = newvalues
                        Next

                    Catch ex As Exception
                        Throw New Exception("Rollen- bzw. Kosten konnten nicht in der Länge angepasst werden" & ex.Message)
                    End Try

                End If

            End If


        Catch ex As Exception
            ' bei einer Projektvorlage gibt es kein Datum - es sollen aber die Werte für Offset und Dauer übernommen werden

            If dauer = 0 And _relEnde > 0 Then


                ' dann sind die Werte initial noch nicht gesetzt worden 
                _startOffsetinDays = DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(_relStart - 1))
                _dauerInDays = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(_relStart - 1), _
                                        StartofCalendar.AddMonths(_relEnde).AddDays(-1)) + 1


            Else
                '  
                _startOffsetinDays = startOffset
                _dauerInDays = dauer

                _relStart = DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset)) + 1
                _relEnde = DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(startOffset + _dauerInDays - 1)) + 1


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

        'Set(value As Integer)

        '    Dim projektStartdate As Date

        '    Try

        '        projektStartdate = Me.Parent.startDate

        '        If value < 0 Then
        '            Throw New ArgumentException("Phase kann nicht vor dem Projekt-Anfang sein")

        '        ElseIf value = 0 And _relStart - 1 > 0 Then
        '            ' wenn z.B nur relstart und relende gesetzt sind 
        '            _startOffsetinDays = DateDiff(DateInterval.Day, projektStartdate, projektStartdate.AddMonths(_relStart - 1))

        '        Else
        '            ' jetzt 
        '            Dim oldlaenge As Integer = _relEnde - _relStart + 1
        '            Dim newlaenge As Integer
        '            _startOffsetinDays = value
        '            _relStart = DateDiff(DateInterval.Month, projektStartdate, projektStartdate.AddDays(value))
        '            _relEnde = DateDiff(DateInterval.Month, projektStartdate, projektStartdate.AddDays(value + _dauerInDays - 1))


        '            newlaenge = _relEnde - _relStart + 1

        '            If newlaenge <> oldlaenge Then
        '                Dim newvalues() As Double
        '                Dim oldvalues() As Double

        '                For r = 1 To Me.CountRoles
        '                    oldvalues = Me.getRole(r).Xwerte
        '                    newvalues = adjustArrayLength(newlaenge, oldvalues, False)
        '                    ' wahrscheionlich muss dafür eine .XwerteReDim Property gemacht werden
        '                    ' die den Redim macht ... 
        '                    Me.getRole(r).Xwerte = newvalues
        '                Next


        '                For k = 1 To Me.CountCosts
        '                    oldvalues = Me.getCost(k).Xwerte
        '                    newvalues = adjustArrayLength(newlaenge, oldvalues, False)
        '                    ' wahrscheionlich muss dafür eine .XwerteReDim Property gemacht werden
        '                    ' die den Redim macht ... 
        '                    Me.getCost(k).Xwerte = newvalues
        '                Next
        '            End If

        '        End If

        '    Catch ex As Exception

        '        ' bei einer Projektvorlage gibt es kein Datum - es soll aber der Wert für den Start-Offset übernommen werden

        '        If value = 0 And (_relStart - 1 > 0) Then
        '            _startOffsetinDays = DateDiff(DateInterval.Day, StartofCalendar, StartofCalendar.AddMonths(relStart - 1)) - 1

        '            ' wenn negativ: zurücksetzen auf NULL
        '            If _startOffsetinDays < 0 Then
        '                _startOffsetinDays = 0
        '            End If

        '        ElseIf value >= 0 Then
        '            _startOffsetinDays = value
        '        Else

        '            Throw New ArgumentException("Phase kann nicht negativen Offset haben, d.h vor dem Projekt beginnen ")

        '        End If


        '    End Try


        'End Set

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

    Public ReadOnly Property Farbe As Object
        Get
            Try
                Farbe = PhaseDefinitions.getPhaseDef(_name).farbe
            Catch ex As Exception
                ' in diesem Fall ist es wahrscheinlich der Name der Projektvorlage 
                Try
                    Farbe = Projektvorlagen.getProject(_name).farbe
                Catch ex1 As Exception
                    Farbe = 0
                    Throw New ArgumentException("Phasen-Name nicht bekannt ...")
                End Try
            End Try
        End Get
    End Property


    Public Property earliestStart As Integer
        Get
            earliestStart = _earliestStart
        End Get
        Set(value As Integer)
            If value >= 0 Then
                If _relStart <> -999 Then
                    If value <= _relStart Then
                        _earliestStart = value
                    Else
                        Throw New ApplicationException("Wert für Earliest Start liegt nach dem aktuellen Start")
                    End If
                Else
                    _earliestStart = value
                End If
            ElseIf value = -999 Then ' die undefiniert Bedingung
                _earliestStart = value
            Else
                Throw New ApplicationException("Wert für Earliest Start kann nicht negativ sein")
            End If

        End Set
    End Property

    Public Property latestStart As Integer
        Get
            latestStart = _latestStart
        End Get
        Set(value As Integer)
            If value >= 0 Then
                If _relStart <> -999 Then
                    If value >= _relStart Then
                        _latestStart = value
                    Else
                        Throw New ApplicationException("Wert für Latest Start liegt vor dem aktuellen Start")
                    End If
                Else
                    _latestStart = value
                End If
            ElseIf value = -999 Then ' die undefiniert Bedingung
                _latestStart = value
            Else
                Throw New ApplicationException("Wert für Earliest Start kann nicht negativ sein")
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
            relStart = _relStart + _Offset
        End Get

        'Set(value As Integer)

        '    If value >= 0 Then

        '        _relStart = value
        '        '_startOffsetinDays = DateDiff(DateInterval.Day, projektStart, projektStart.AddMonths(value))

        '    Else

        '        Throw New ApplicationException("Phasen-Start kann nicht negativ sein ..")

        '    End If

        'End Set
        'Set(value As Integer)
        '    If value >= 0 Then
        '        If _earliestStart <> -999 And _latestStart <> -999 Then
        '            If value + _Offset >= _earliestStart And value + _Offset <= _latestStart Then
        '                If _relEnde <> -999 Then
        '                    If value <= _relEnde Then
        '                        _relStart = value
        '                    Else
        '                        Throw New ApplicationException("Start liegt nach dem Ende")
        '                    End If
        '                    _relStart = value
        '                End If
        '            Else
        '                Throw New ApplicationException("Start liegt ausserhalb des zugelassenen Korridors")
        '            End If
        '        Else
        '            _relStart = value
        '        End If
        '    Else
        '        Throw New ApplicationException("Phasen-Start kann nicht negativ sein ..")
        '    End If

        'End Set
    End Property

    'Public WriteOnly Property relEnde(projektStartDate As Date) As Integer

    '    Set(value As Integer)

    '        If value >= _relStart Then

    '            _relEnde = value
    '            _dauerInDays = DateDiff(DateInterval.Day, projektStartDate, projektStartDate.AddMonths(value))

    '        Else

    '            Throw New ApplicationException("Phasen-Start kann nicht negativ sein ..")

    '        End If

    '    End Set
    'End Property

    Public ReadOnly Property relEnde As Integer
        Get
            relEnde = _relEnde + _Offset
        End Get

        'Set(value As Integer)

        '    If value >= _relStart Then

        '        _relEnde = value
        '        '_dauerInDays = DateDiff(DateInterval.Day, projektStartDate, projektStartDate.AddMonths(value))

        '    Else

        '        Throw New ApplicationException("Phasen-Start kann nicht negativ sein ..")

        '    End If

        'End Set
        'Set(value As Integer)
        '    If value >= 0 Then
        '        If value >= _relStart Then
        '            _relEnde = value
        '        Else
        '            Throw New ApplicationException("das Ende kann nicht vor dem Start sein ")
        '        End If
        '    Else
        '        Throw New ApplicationException("das Ende kann nicht negativ sein ")
        '    End If

        'End Set
    End Property

    Public Property name As String
        Get
            name = _name
        End Get
        Set(value As String)
            If Len(value) > 1 Then
                _name = value
            Else
                Throw New ApplicationException("Name muss mindestens zwei Zeichen lang sein ...")
            End If

        End Set
    End Property


    Public Sub calculateLineCoord(ByVal zeile As Integer, ByVal nummer As Integer, ByVal gesamtZahl As Integer, _
                                  ByRef top1 As Double, ByRef left1 As Double, ByRef top2 As Double, ByRef left2 As Double, ByVal linienDicke As Double)

        Try

            Dim projektStartdate As Date = Me.Parent.startDate
            
            Dim korrPosition As Double = nummer / gesamtZahl
            Dim faktor As Double = linienDicke / boxHeight
            Dim startpunkt As Integer = DateDiff(DateInterval.Day, StartofCalendar, projektStartdate)

            If startpunkt < 0 Then
                Throw New Exception("calculate Line Coord: Projektstart liegt vor Start of Calendar ...")
            End If

            Dim phasenStart As Integer = startpunkt + Me.startOffsetinDays
            Dim phasenDauer As Integer = Me.dauerInDays

            ' absolute Setzung - dadurch wird verhindert, daß die Linien sehr schmal gezeichnet werden ... 
            ' es soll immer gleich groß gezeichnet werden - einfach überschreiben - das ist rvtl besser;
            ' das muss einfach noch herausgefunden werden 
            gesamtZahl = 1
            nummer = 1


            If gesamtZahl <= 0 Then
                Throw New ArgumentException("unzulässige Gesamtzahl" & gesamtZahl)
            End If

            ' korrigiere, aber breche nicht ab wenn die Nummer der Line größer als die Gesamtzahl ist ... 
            If nummer > gesamtZahl Then
                nummer = gesamtZahl
            End If

            ' ausrechnen des Korrekturfaktors

            korrPosition = nummer / (gesamtZahl + 1)
            

            If phasenStart >= 0 And phasenDauer > 0 Then

                ' das folgende ist mühsam ausprobiert - um die Linien in unterschiedicher Stärke in der Projekt Form zu platzieren - möglichst auch jeweils mittig
                If gesamtZahl <= 3 Then
                    top1 = topOfMagicBoard + (zeile - 0.95) * boxHeight + korrPosition * boxHeight - linienDicke / 2
                Else
                    top1 = topOfMagicBoard + (zeile - 1.06) * boxHeight + korrPosition * boxHeight - linienDicke / 2
                End If

                top2 = top1

                left1 = (phasenStart / 365) * boxWidth * 12
                left2 = ((phasenStart + phasenDauer) / 365) * boxWidth * 12

            Else
                Throw New ArgumentException("es kann keine Line berechnet werden für : " & Me.name)
            End If

        Catch ex As Exception
            Throw New ArgumentException("es kann keine Line berechnet werden für : " & Me.name)
        End Try



    End Sub
    Public Sub AddRole(ByVal role As clsRolle)

        AllRoles.Add(role)

    End Sub

    Public Sub AddResult(ByVal result As clsResult)

        AllResults.Add(result)

    End Sub

    Public ReadOnly Property RollenListe() As List(Of clsRolle)

        Get
            RollenListe = AllRoles
        End Get

    End Property

    Public ReadOnly Property ResultListe() As List(Of clsResult)

        Get
            ResultListe = AllResults
        End Get

    End Property

    Public ReadOnly Property KostenListe() As List(Of clsKostenart)

        Get
            KostenListe = AllCosts
        End Get

    End Property


    Public ReadOnly Property CountRoles() As Integer

        Get
            CountRoles = AllRoles.Count
        End Get

    End Property

    Public ReadOnly Property CountResults() As Integer

        Get
            CountResults = AllResults.Count
        End Get

    End Property

    Public ReadOnly Property DauerM() As Integer

        Get
            DauerM = _relEnde - _relStart + 1
        End Get

    End Property


    Public Sub CopyTo(ByRef newphase As clsPhase)
        Dim r As Integer, k As Integer
        Dim newrole As clsRolle
        Dim newcost As clsKostenart
        Dim newresult As clsResult

        With newphase
            .minDauer = Me._minDauer
            .maxDauer = Me._maxDauer
            .earliestStart = Me._earliestStart
            .latestStart = Me._latestStart
            .Offset = Me._Offset

           
            .changeStartandDauer(Me._startOffsetinDays, Me._dauerInDays)


            .name = _name

            For r = 1 To Me.CountRoles
                newrole = New clsRolle(relEnde - relStart)
                Me.getRole(r).CopyTo(newrole)
                .AddRole(newrole)
            Next r

            For r = 1 To Me.AllResults.Count
                newresult = New clsResult(parent:=newphase)
                Me.getResult(r).CopyTo(newresult)
                .AddResult(newresult)
            Next

            For k = 1 To Me.CountCosts
                newcost = New clsKostenart(relEnde - relStart)
                Me.getCost(k).CopyTo(newcost)
                .AddCost(newcost)
            Next k

        End With

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

    Public ReadOnly Property getResult(ByVal index As Integer) As clsResult

        Get
            getResult = AllResults.Item(index - 1)
        End Get

    End Property

    Public ReadOnly Property getResult(ByVal key As String) As clsResult

        Get
            Dim tmpResult As clsResult = Nothing
            Dim found As Boolean = False
            Dim r As Integer = 1

            While r <= Me.CountResults And Not found

                If AllResults.Item(r - 1).name = key Then
                    found = True
                    tmpResult = AllResults.Item(r - 1)
                Else
                    r = r + 1
                End If

            End While

            getResult = tmpResult

            If Not found Then
                Throw New ArgumentException("Result " & key & " nicht gefunden ")
            End If


        End Get

    End Property

    Public Sub AddCost(ByVal cost As clsKostenart)

        AllCosts.Add(cost)

    End Sub


    Public ReadOnly Property CountCosts() As Integer

        Get
            CountCosts = AllCosts.Count
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

    Public Sub New(ByRef parent As clsProjekt)

        AllRoles = New List(Of clsRolle)
        AllCosts = New List(Of clsKostenart)
        AllResults = New List(Of clsResult)
        _minDauer = 1
        _maxDauer = 60
        _Offset = 0
        _earliestStart = -999
        _latestStart = -999
        _Parent = parent
        _vorlagenParent = Nothing


    End Sub

    Public Sub New(ByRef parent As clsProjektvorlage, ByVal isVorlage As Boolean)
        ' Variable isVorlage dient lediglich dazu, eine weitere Signatur für einen Konstruktor zu bekommen 
        ' dieser Konstruktor wird für parent = Vorlage benutzt 


        AllRoles = New List(Of clsRolle)
        AllCosts = New List(Of clsKostenart)
        AllResults = New List(Of clsResult)
        _minDauer = 1
        _maxDauer = 60
        _Offset = 0
        _earliestStart = -999
        _latestStart = -999
        _Parent = Nothing
        _vorlagenParent = parent



    End Sub

    'Public Sub New()

    '    AllRoles = New List(Of clsRolle)
    '    AllCosts = New List(Of clsKostenart)
    '    AllResults = New List(Of clsResult)
    '    _minDauer = 1
    '    _maxDauer = 60
    '    _Offset = 0
    '    _earliestStart = -999
    '    _latestStart = -999


    'End Sub

End Class
