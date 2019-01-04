Public Class clsProjektHistorie

    ' Methoden für Behandlung der Projekthistorie 
    ' die Projekthistorie ist eine aufsteigend nach dem Datum sortierte Liste all der Planungs-Stände
    ' eines bestimmten Projektes 
    ' _currentIndex ist ein Zeiger auf das Element der Historie, das zuletzt bearbeitet wurd
    ' _currentIndex ist insbesondere für prevdiff und nextdiff wichtig: Methoden, die auf das nächste Element "springen", 
    ' das sich im angegebenen Kriterium vom ElementAt(_currentIndex) unterscheiden 

    ' die _liste enthält die Timestamps des Projektes
    Private _liste As SortedList(Of Date, clsProjekt)

    ' 3.1.19 die _pfvliste enthält die Timestamps der Vorgabe-Projekte durch den Portfolio Manager 
    Private _pfvliste As SortedList(Of Date, clsProjekt)

    Private _currentIndex As Integer

    Public Property liste As SortedList(Of Date, clsProjekt)
        Get
            liste = _liste
        End Get
        Set(value As SortedList(Of Date, clsProjekt))
            _liste = value
        End Set
    End Property

    Public ReadOnly Property pfvListe As SortedList(Of Date, clsProjekt)
        Get
            pfvListe = _pfvliste
        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob das angegebene Datum in der Projekt-Historie existiert ... 
    ''' </summary>
    ''' <param name="dateItem"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property contains(ByVal dateItem As Date) As Boolean
        Get
            contains = _liste.ContainsKey(dateItem)
        End Get
    End Property

    ''' <summary>
    ''' entfernt das ELement mit Datum dateItem 
    ''' wenn es nicht existiert, wird eine Exception geworfen ... 
    ''' </summary>
    ''' <param name="dateItem"></param>
    ''' <remarks></remarks>
    Public Sub remove(ByVal dateItem As Date)

        _liste.Remove(dateItem)

    End Sub

    ''' <summary>
    ''' gibt das Element zurück, das tsDate as Schlüssel hat 
    ''' </summary>
    ''' <param name="tsDate"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property item(ByVal tsDate As Date) As clsProjekt
        Get
            If _liste.ContainsKey(tsDate) Then
                item = _liste.Item(tsDate)
            Else
                item = Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property getZeitraum As String
        Get

            If _liste.Count > 1 Then
                getZeitraum = _liste.First.Value.timeStamp.ToShortDateString & " - " & _
                              _liste.Last.Value.timeStamp.ToShortDateString
            Else
                If _liste.Count > 0 Then
                    getZeitraum = _liste.First.Value.timeStamp.ToShortDateString & " - " & _
                              _liste.First.Value.timeStamp.ToShortDateString
                Else
                    getZeitraum = "keine Historie vorhanden ..."
                End If
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl an Historien Elementen zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count As Integer
        Get
            Count = _liste.Count
        End Get
    End Property


    ''' <summary>
    ''' gibt das Element zurück , das die erste Vorgabe vom Portfolio Manager war
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property beauftragung As clsProjekt
        Get

            If _pfvliste.Count > 0 Then
                beauftragung = _pfvliste.First.Value
            Else
                beauftragung = Nothing
            End If


        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück, das zum Zeitpunkt refDate das vom Portfolio Manager zuletzt beauftragte war 
    ''' Nothing, wenn zu dem Zeitpunkt noch keine Beauftragung existierte 
    ''' </summary>
    ''' <param name="refDate"></param>
    ''' <returns></returns>
    Public ReadOnly Property lastBeauftragung(ByVal refDate As Date) As clsProjekt

        Get
            Dim found = False
            Dim tmpResult As clsProjekt = Nothing

            If _pfvliste.Count = 0 Then
                ' nichts tun, es gibt keine Beauftragung 
            Else
                Dim ix As Integer = _pfvliste.Count - 1
                Dim curTimestamp As Date = _pfvliste.ElementAt(ix).Value.timeStamp

                found = (refDate > curTimestamp)
                Do While ix >= 0 And Not found
                    ix = ix - 1
                    curTimestamp = _pfvliste.ElementAt(ix).Value.timeStamp
                    found = (refDate > curTimestamp)
                Loop

                If found Then
                    tmpResult = _pfvliste.ElementAt(ix).Value
                End If

            End If

            lastBeauftragung = tmpResult

        End Get
    End Property


    ''' <summary>
    ''' löscht die Historie
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clear()

        _liste.Clear()

    End Sub

    ''' <summary>
    ''' gibt für den Aufbau einer Milestone Trendanalyse einen Array mit den Plan-Daten eines ausgewählten Meilensteins zurück
    ''' der array hat die Dimension (Start-Monat des Projektes/Aufzeichnungs-Start) bis (aktueller Monat)
    ''' wenn es für einen bestimmten Monat keine Werte gibt, dann wird der Vormonats Wert genommen 
    ''' </summary>
    ''' <param name="milestoneName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMtaDates(ByVal milestoneName As String, ByVal von As Integer, ByVal bis As Integer) As Date()
        Get
            Dim tmpValues As Date()
            Dim heute As Date = Date.Now
            Dim milestoneDate As Date
            Dim oldIndex As Integer = _currentIndex
            Dim laenge As Integer
            Dim currentproj As clsProjekt



            Try

                ' bestimme die Dimension; es ist bereits sichergestellt, daß laenge > 0 ist 
                laenge = bis - von

                ReDim tmpValues(laenge)

                Dim tmpDate As Date
                For i = 0 To laenge
                    ' StartofCalendar ist 00:00 Uhr, deswegen kommt man zum Vortag , 23:55 , indem man 5 Min abzieht 
                    tmpDate = StartofCalendar.Date.AddMonths(von + i).AddMinutes(-5)

                    Try
                        currentproj = Me.ElementAtorBefore(tmpDate)

                    Catch ex As Exception
                        currentproj = Nothing
                    End Try


                    If currentproj Is Nothing Then

                        milestoneDate = awinSettings.nullDatum

                    ElseIf DateDiff(DateInterval.Month, tmpDate, currentproj.timeStamp) = 0 Then
                        ' in diesem Fall wurde ein Planungs-Stand im gesuchten Monat gefunden ...
                        
                        milestoneDate = currentproj.getMilestoneDate(milestoneName)
                        If DateDiff(DateInterval.Day, StartofCalendar, milestoneDate) < 0 Then
                            milestoneDate = awinSettings.nullDatum
                        End If

                    Else
                        ' in diesem Fall wurde kein Planungs-Stand im gesuchten Monat gefunden ...
                        If i > 0 Then
                            ' jetzt wird gekennzeichnet, dass einfach die Bewertung des Vormonats übernommen wurde: + 12 Std
                            ' aber nur, wenn der Wert von vorher nicht schon die Kennung hatte ....
                            If DateDiff(DateInterval.Hour, tmpValues(i - 1).Date, tmpValues(i - 1)) < 7 Then
                                milestoneDate = tmpValues(i - 1).AddHours(12)
                            End If

                        Else
                            ' damit wird gekennzeichnet, dass es eigentlich keinen Wert im Berichtsmonat gab
                            milestoneDate = currentproj.getMilestoneDate(milestoneName).AddHours(12)
                            'milestoneDate = awinSettings.nullDatum
                        End If

                    End If


                    tmpValues(i) = milestoneDate

                Next

                _currentIndex = oldIndex
                getMtaDates = tmpValues

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try


        End Get
    End Property

    ''' <summary>
    ''' holt das Item aus der History, das vor dem aktuellen liegt und im ChangeCriteria einen anderen Wert hat
    ''' </summary>
    ''' <param name="changeCriteria"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PrevDiff(ByVal changeCriteria As Integer) As clsProjekt
        Get

            Dim currentValues() As Double
            Dim tstValues() As Double
            Dim currentproj As clsProjekt
            Dim tstproj As clsProjekt
            Dim tstindex As Integer = _currentIndex - 1
            Dim found As Boolean

            Try
                currentproj = _liste.ElementAt(_currentIndex).Value
                tstproj = _liste.ElementAt(tstindex).Value
            Catch ex As Exception
                Throw New ArgumentException("ungültiger Wert für _currentIndex: " & _currentIndex)
            End Try


            Select Case changeCriteria

                Case PThcc.perscost

                    currentValues = currentproj.getAllPersonalKosten
                    tstValues = tstproj.getAllPersonalKosten

                Case PThcc.othercost

                    currentValues = currentproj.getGesamtAndereKosten
                    tstValues = tstproj.getGesamtAndereKosten

                Case PThcc.budget

                    ReDim currentValues(0)
                    ReDim tstValues(0)
                    currentValues(0) = currentproj.Erloes
                    tstValues(0) = tstproj.Erloes

                Case PThcc.ergebnis

                    ReDim currentValues(0)
                    ReDim tstValues(0)
                    currentValues(0) = currentproj.Erloes - currentproj.getSummeKosten
                    tstValues(0) = tstproj.Erloes - tstproj.getSummeKosten

                Case PThcc.fitrisk

                    ReDim currentValues(1)
                    ReDim tstValues(1)
                    currentValues(0) = currentproj.StrategicFit
                    currentValues(1) = currentproj.Risiko
                    tstValues(0) = tstproj.StrategicFit
                    tstValues(1) = tstproj.Risiko

                Case PThcc.resultdates

                    Throw New ArgumentException("wird noch nicht unterstützt")

                Case PThcc.projektampel

                    ReDim currentValues(0)
                    ReDim tstValues(0)
                    currentValues(0) = currentproj.ampelStatus
                    tstValues(0) = tstproj.ampelStatus

                Case PThcc.resultampel

                    Throw New ArgumentException("wird noch nicht unterstützt")

                    'Dim ix = 0
                    'For p = 1 To currentproj.CountPhases
                    '    Dim cphase As clsPhase = currentproj.getPhase(p)
                    '    For r = 1 To cphase.CountResults

                    '    Next
                    'Next

                Case PThcc.phasen

                    currentValues = currentproj.getPhaseInfos
                    tstValues = tstproj.getPhaseInfos


                Case Else
                    Throw New ArgumentException("kein gültiges Kriterium angegeben: " & changeCriteria)
            End Select


            found = arraysAreDifferent(currentValues, tstValues)

            Do While Not found And tstindex > 0

                tstindex = tstindex - 1
                tstproj = _liste.ElementAt(tstindex).Value

                Select Case changeCriteria

                    Case PThcc.perscost

                        tstValues = tstproj.getAllPersonalKosten

                    Case PThcc.othercost

                        tstValues = tstproj.getGesamtAndereKosten

                    Case PThcc.budget

                        tstValues(0) = tstproj.Erloes

                    Case PThcc.ergebnis

                        tstValues(0) = tstproj.Erloes - tstproj.getSummeKosten

                    Case PThcc.fitrisk

                        tstValues(0) = tstproj.StrategicFit
                        tstValues(1) = tstproj.Risiko

                    Case PThcc.resultdates

                        Throw New ArgumentException("wird noch nicht unterstützt")

                    Case PThcc.projektampel

                        tstValues(0) = tstproj.ampelStatus

                    Case PThcc.resultampel

                        Throw New ArgumentException("wird noch nicht unterstützt")

                    Case PThcc.phasen

                        tstValues = tstproj.getPhaseInfos

                    Case Else
                        Throw New ArgumentException("kein gültiges Kriterium angegeben: " & changeCriteria)
                End Select


                found = arraysAreDifferent(currentValues, tstValues)

            Loop



            If found Then
                _currentIndex = tstindex
                PrevDiff = tstproj
            Else
                Throw New ArgumentException("kein Element gefunden")
            End If

        End Get
    End Property

    ''' <summary>
    ''' holt das Item aus der History, das nach dem aktuellen liegt und im ChangeCriteria einen anderen Wert hat
    ''' </summary>
    ''' <param name="changeCriteria"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NextDiff(ByVal changeCriteria As Integer) As clsProjekt
        Get

            Dim currentValues() As Double
            Dim tstValues() As Double
            Dim currentproj As clsProjekt
            Dim tstproj As clsProjekt
            Dim tstindex As Integer = _currentIndex + 1
            Dim maxIndex = _liste.Count - 1
            Dim found As Boolean

            Try
                currentproj = _liste.ElementAt(_currentIndex).Value
                tstproj = _liste.ElementAt(tstindex).Value
            Catch ex As Exception
                Throw New ArgumentException("ungültiger Wert für _currentIndex: " & _currentIndex)
            End Try


            Select Case changeCriteria

                Case PThcc.perscost

                    currentValues = currentproj.getAllPersonalKosten
                    tstValues = tstproj.getAllPersonalKosten

                Case PThcc.othercost

                    currentValues = currentproj.getGesamtAndereKosten
                    tstValues = tstproj.getGesamtAndereKosten

                Case PThcc.budget

                    ReDim currentValues(0)
                    ReDim tstValues(0)
                    currentValues(0) = currentproj.Erloes
                    tstValues(0) = tstproj.Erloes

                Case PThcc.ergebnis

                    ReDim currentValues(0)
                    ReDim tstValues(0)
                    currentValues(0) = currentproj.Erloes - currentproj.getSummeKosten
                    tstValues(0) = tstproj.Erloes - tstproj.getSummeKosten

                Case PThcc.fitrisk

                    ReDim currentValues(1)
                    ReDim tstValues(1)
                    currentValues(0) = currentproj.StrategicFit
                    currentValues(1) = currentproj.Risiko
                    tstValues(0) = tstproj.StrategicFit
                    tstValues(1) = tstproj.Risiko

                Case PThcc.resultdates

                    Throw New ArgumentException("wird noch nicht unterstützt")

                Case PThcc.projektampel

                    ReDim currentValues(0)
                    ReDim tstValues(0)
                    currentValues(0) = currentproj.ampelStatus
                    tstValues(0) = tstproj.ampelStatus

                Case PThcc.resultampel

                    Throw New ArgumentException("wird noch nicht unterstützt")

                Case PThcc.phasen

                    currentValues = currentproj.getPhaseInfos
                    tstValues = tstproj.getPhaseInfos

                Case Else
                    Throw New ArgumentException("kein gültiges Kriterium angegeben: " & changeCriteria)
            End Select


            found = arraysAreDifferent(currentValues, tstValues)

            Do While Not found And tstindex < maxIndex

                tstindex = tstindex + 1
                tstproj = _liste.ElementAt(tstindex).Value

                Select Case changeCriteria

                    Case PThcc.perscost

                        tstValues = tstproj.getAllPersonalKosten

                    Case PThcc.othercost

                        tstValues = tstproj.getGesamtAndereKosten

                    Case PThcc.budget

                        tstValues(0) = tstproj.Erloes

                    Case PThcc.ergebnis

                        tstValues(0) = tstproj.Erloes - tstproj.getSummeKosten

                    Case PThcc.fitrisk

                        tstValues(0) = tstproj.StrategicFit
                        tstValues(1) = tstproj.Risiko

                    Case PThcc.resultdates

                        Throw New ArgumentException("wird noch nicht unterstützt")

                    Case PThcc.projektampel

                        tstValues(0) = tstproj.ampelStatus

                    Case PThcc.resultampel

                        Throw New ArgumentException("wird noch nicht unterstützt")

                    Case PThcc.phasen

                        tstValues = tstproj.getPhaseInfos

                    Case Else
                        Throw New ArgumentException("kein gültiges Kriterium angegeben: " & changeCriteria)
                End Select


                found = arraysAreDifferent(currentValues, tstValues)

            Loop



            If found Then
                _currentIndex = tstindex
                NextDiff = tstproj
            Else
                Throw New ArgumentException("kein Element gefunden")
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt das erste Element der Historie zurück 
    ''' setzt _currentIndex auf dieses Element 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property First As clsProjekt
        Get
            If _liste.Count > 0 Then
                First = _liste.First.Value
            Else
                First = Nothing
            End If
            _currentIndex = 0
        End Get
    End Property

    ''' <summary>
    ''' gibt das letzte Element der Historie zurück
    ''' setzt _currentIndex auf dieses Element 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Last As clsProjekt
        Get

            If _liste.Count > 0 Then
                Last = _liste.Last.Value
                _currentIndex = _liste.Count - 1
            Else
                Last = Nothing
                _currentIndex = 0
            End If
            

        End Get
    End Property

    ''' <summary>
    ''' gibt das Projekt aus der Historie an der Position index zurück; 
    ''' setzt _currentIndex auf dieses Element  
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ElementAt(ByVal index As Integer) As clsProjekt
        Get
            If index < 0 Or index > _liste.Count - 1 Then
                ElementAt = Nothing
                ' alt tk, 14.11.16
                'Throw New ArgumentException("index liegt ausserhalb der zulässigen Grenzen")
            Else
                ElementAt = _liste.ElementAt(index).Value
                _currentIndex = index
            End If
        End Get
    End Property

    ''' <summary>
    ''' löscht das Element an der Position index aus der Historie
    ''' lässt den Index unverändert 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <remarks></remarks>
    Public Sub RemoveAt(ByVal index As Integer)

        If index < 0 Or index > _liste.Count - 1 Then
            Throw New ArgumentException("index liegt ausserhalb der zulässigen Grenzen")
        Else
            _liste.RemoveAt(index)
            If _currentIndex >= index And _currentIndex > 0 Then
                _currentIndex = _currentIndex - 1
            End If
        End If

    End Sub

    ''' <summary>
    ''' gibt das Element zurück, das als letztes vor dem angegebenen Datum liegt  
    ''' </summary>
    ''' <param name="suchDatum"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ElementAtorBefore(ByVal suchDatum As Date) As clsProjekt

        Get

            ' die Ausschlusskriterien vorher prüfen
            If suchDatum < _liste.First.Key Then

                If DateDiff(DateInterval.Second, _liste.First.Key, suchDatum) = 0 Then
                    _currentIndex = 0
                    ElementAtorBefore = _liste.First.Value
                Else
                    ElementAtorBefore = Nothing
                End If

            ElseIf suchDatum > _liste.Last.Key Then

                _currentIndex = _liste.Count - 1
                ElementAtorBefore = _liste.Last.Value

            Else
                ' das Suchdatum liegt zwischen erstem und letztem Element
                Dim found As Boolean = False

                Dim i As Integer = 0

                While i <= _liste.Count - 1 And Not found

                    If DateDiff(DateInterval.Second, _liste.ElementAt(i).Key, suchDatum) = 0 Then
                        found = True
                    ElseIf suchDatum > _liste.ElementAt(i).Key Then
                        i = i + 1
                    Else
                        ' es gibt keinen exakten Match, aber das Abbruch Kriterium ist erfüllt 
                        found = True
                        If i > 0 Then
                            i = i - 1
                        Else
                            i = 0
                        End If

                    End If

                End While


                If found Then
                    _currentIndex = i
                Else
                    _currentIndex = _liste.Count - 1
                End If

                ElementAtorBefore = _liste.ElementAt(_currentIndex).Value

            End If



        End Get
    End Property
    ''' <summary>
    ''' fügt der pfvListe ein Projekt hinzu
    ''' </summary>
    ''' <param name="value"></param>
    Public Sub AddPfv(ByVal value As clsProjekt)

        Try
            Dim ok As Boolean = True

            If _pfvliste.ContainsKey(value.timeStamp) Then
                ok = _pfvliste.Remove(value.timeStamp)
            End If

            If ok Then
                _pfvliste.Add(value.timeStamp, value)
            End If

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' fügt ein Element der Historie hinzu 
    ''' </summary>
    ''' <param name="datum"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal datum As Date, ByVal value As clsProjekt)

        Try

            If _liste.Count > 0 Then
                If _liste.First.Value.name <> value.name Then
                    Throw New ArgumentException _
                        ("Projekte mit unterschiedlichen Namen können nicht in einer Projekt-Historie sein")
                Else
                    If _liste.ContainsKey(datum) Then
                        _liste.Remove(datum)
                    End If
                    _liste.Add(datum, value)
                End If
            Else
                _liste.Add(datum, value)
            End If


        Catch ex As Exception
            Throw New ArgumentException("es gibt keine Einträge in der Datenbank")
        End Try

    End Sub

    ''' <summary>
    ''' setzt den _currentIndex auf das entsprechende Element
    ''' oder gibt den Index des aktuellen Elements zurück  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property currentIndex As Integer

        Get
            currentIndex = _currentIndex
        End Get

        Set(value As Integer)

            If value >= 0 And value <= _liste.Count - 1 Then
                _currentIndex = value
            Else
                Throw New ArgumentException("Wert ausserhalb der zulässigen Grenze: " & value)
            End If

        End Set

    End Property

    Sub New()
        _liste = New SortedList(Of Date, clsProjekt)
        _pfvliste = New SortedList(Of Date, clsProjekt)
        _currentIndex = -1
    End Sub

End Class
