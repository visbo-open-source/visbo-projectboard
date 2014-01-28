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
                '_dauerInDays = DateDiff(DateInterval.Day, projektStartdate.AddMonths(_relStart - 1), _
                '                        projektStartdate.AddMonths(_relEnde).AddDays(-1)) + 1
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


            ElseIf dauer = 0 And _relEnde = 0 Then

                Throw New ArgumentException("Phase kann nicht Dauer = 0 haben ")

            Else
                '  

                _startOffsetinDays = startOffset
                _dauerInDays = dauer

                Dim oldlaenge As Integer = _relEnde - _relStart + 1
                Dim newlaenge As Integer


                phaseStartdate = Me.getStartDate
                phaseEndDate = Me.getEndDate



                _relStart = getColumnOfDate(phaseStartdate) - projektstartColumn + 1
                _relEnde = getColumnOfDate(phaseEndDate) - projektstartColumn + 1


                If awinSettings.autoCorrectBedarfe Then

                    newlaenge = _relEnde - _relStart + 1

                    Dim newvalues() As Double
                    Dim oldvalues() As Double


                    Try

                        For r = 1 To Me.CountRoles
                            oldvalues = Me.getRole(r).Xwerte
                            oldlaenge = oldvalues.Length
                            If newlaenge <> oldlaenge Then
                                newvalues = adjustArrayLength(newlaenge, oldvalues, False)
                                Me.getRole(r).Xwerte = newvalues
                            End If
                            
                        Next


                        For k = 1 To Me.CountCosts
                            oldvalues = Me.getCost(k).Xwerte
                            oldlaenge = oldvalues.Length
                            If newlaenge <> oldlaenge Then
                                newvalues = adjustArrayLength(newlaenge, oldvalues, False)
                                Me.getCost(k).Xwerte = newvalues
                            End If
                            
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
                '_dauerInDays = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(_relStart - 1), _
                '                        StartofCalendar.AddMonths(_relEnde).AddDays(-1)) + 1
                _dauerInDays = calcDauerIndays(projektStartdate.AddDays(_startOffsetinDays), _relEnde - _relStart + 1, True)


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
                Throw New Exception("Dauer muss mindestens 1 Tag sein ...")
            End If

        End Get

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
            'Dim checkValue As Integer = _relEnde + _Offset

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

        ' in Abhängigkeit von milestoneFreeFloat: 
        ' es wird geprüft, ob der Meilenstein innerhalb der Projektgrenzen ist 
        ' wenn nein , wird entweder auf Projektstart gesetzt, wenn er vor dem Projektstart liegt 
        ' oder auf Projektende, wenn er nach dem Projektende liegt 

        If awinSettings.milestoneFreeFloat Then
            ' nichts verändern ....
        ElseIf IsNothing(_vorlagenParent) Then
            If result.offset + Me.startOffsetinDays > Me.Parent.dauerInDays - 1 Then
                'result.offset = result.offset - (result.offset + Me.startOffsetinDays - (Me.Parent.dauerInDays - 1))
                result.offset = Me.Parent.dauerInDays - 1 - Me.startOffsetinDays
            ElseIf result.offset + Me.startOffsetinDays < 0 Then
                result.offset = -1 * Me.startOffsetinDays
            End If
        Else
            If result.offset + Me.startOffsetinDays > Me.VorlagenParent.dauerInDays - 1 Then
                'result.offset = result.offset - (result.offset + Me.startOffsetinDays - (Me.Parent.dauerInDays - 1))
                result.offset = Me.VorlagenParent.dauerInDays - 1 - Me.startOffsetinDays
            ElseIf result.offset + Me.startOffsetinDays < 0 Then
                result.offset = -1 * Me.startOffsetinDays
            End If
        End If

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



    Public Sub CopyTo(ByRef newphase As clsPhase)
        Dim r As Integer, k As Integer
        Dim newrole As clsRolle
        Dim newcost As clsKostenart
        Dim newresult As clsResult
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

           

            .name = _name

            For r = 1 To Me.CountRoles
                'newrole = New clsRolle(relEnde - relStart)

                dimension = Me.getRole(r).getDimension
                newrole = New clsRolle(dimension)
                Me.getRole(r).CopyTo(newrole)
                .AddRole(newrole)
            Next r


            For k = 1 To Me.CountCosts
                'newcost = New clsKostenart(relEnde - relStart)

                dimension = Me.getCost(k).getDimension
                newcost = New clsKostenart(dimension)
                Me.getCost(k).CopyTo(newcost)
                .AddCost(newcost)
            Next k


            ' Änderung 16.1.2014: zuerst die Rollen und Kosten übertragen, dann die relStart und RelEnde, dann die Results
            ' die evtrl enstehende Inkonsistenz zwischen Länder der Arrays der Rollen/Kostenarten und dem neuen relende/relstart wird in Kauf genommen 
            ' und nur korrigiert , wenn explizit gewünscht (Parameter awinsettings.autoCorrectBedarfe = true 
            
            .changeStartandDauer(Me._startOffsetinDays, Me._dauerInDays)

            For r = 1 To Me.AllResults.Count
                newresult = New clsResult(parent:=newphase)
                Me.getResult(r).CopyTo(newresult)
                .AddResult(newresult)
            Next

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

    ''' <summary>
    ''' gibt das Objekt Meilenstein mit dem angegebenen NAmen zurück. 
    ''' Wenn der Meilenstein nicht existiert, wird Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="key">Name des Meilensteines</param>
    ''' <value></value>
    ''' <returns>Objekt vom Typ Result</returns>
    ''' <remarks>
    ''' Rückgabe von Nothing ist schneller als über Throw Exception zu arbeiten</remarks>
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

    Public ReadOnly Property VorlagenParent() As clsProjektvorlage
        Get
            VorlagenParent = _vorlagenParent
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
