Public Class clsProjektHistorie

    ' Methoden für Behandlung der Projekthistorie 
    ' die Projekthistorie ist eine aufsteigend nach dem Datum sortierte Liste all der Planungs-Stände
    ' eines bestimmten Projektes 
    ' _currentIndex ist ein Zeiger auf das Element der Historie, das zuletzt bearbeitet wurd
    ' _currentIndex ist insbesondere für prevdiff und nextdiff wichtig: Methoden, die auf das nächste Element "springen", 
    ' das sich im angegebenen Kriterium vom ElementAt(_currentIndex) unterscheiden 

    Private _liste As New SortedList(Of Date, clsProjekt)
    Private _currentIndex As Integer

    Public Property liste As SortedList(Of Date, clsProjekt)
        Get
            liste = _liste
        End Get
        Set(value As SortedList(Of Date, clsProjekt))
            _liste = value
        End Set
    End Property

    Public ReadOnly Property getZeitraum As String
        Get

            If _liste.Count > 1 Then
                getZeitraum = _liste.First.Value.timeStamp.ToShortDateString & " - " & _
                              _liste.Last.Value.timeStamp.ToShortDateString
            Else
                Throw New ArgumentException("Historie enthält weniger als 2 Einträge")
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
    ''' gibt das Projekt zurück , das relativ zur aktuellen Position den status "freigegeben/beauftragt" hat 
    ''' falls es das nicht gibt, wird eine Exception geworfen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property letzteFreigabe As clsProjekt
        Get

            Dim found As Boolean = False
            Dim index As Integer = _currentIndex - 1


            ' jetzt wird der Planungs-Stand der letzten Freigabe gesucht 
            Do While index >= 0 And Not found

                If _liste.ElementAt(index).Value.Status = ProjektStatus(1) Then
                    found = True
                Else
                    index = index - 1
                End If

            Loop

            If Not found Then
                Throw New ArgumentException("es gibt keinen letzten Freigabe Stand")
            Else
                _currentIndex = index
                letzteFreigabe = _liste.ElementAt(index).Value
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück , das das erste Mal den Status "Beauftragung/freigegeben" hat
    ''' setzt den Index auf dieses Element
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property beauftragung As clsProjekt
        Get

            Dim abbruch As Boolean = False
            Dim index As Integer = 0
            Dim anzSnapshots As Integer = _liste.Count - 1


            ' jetzt wird der Planungs-Stand der Beauftragung gesucht 
            Do While _liste.ElementAt(index).Value.Status <> ProjektStatus(1) And _
                     _liste.ElementAt(index).Value.Status <> ProjektStatus(2) And Not abbruch
                If index + 1 < anzSnapshots Then
                    index = index + 1
                Else
                    abbruch = True
                End If
            Loop

            If abbruch Then
                Throw New ArgumentException("es gibt keine Beauftragung")
            Else
                _currentIndex = index
                beauftragung = _liste.ElementAt(index).Value
            End If

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
                    tmpDate = StartofCalendar.AddMonths(von + i).AddMinutes(-5)

                    Try
                        currentproj = Me.ElementAtorBefore(tmpDate)

                    Catch ex As Exception
                        currentproj = Nothing
                    End Try


                    If currentproj Is Nothing Then

                        milestoneDate = awinSettings.nullDatum

                    ElseIf DateDiff(DateInterval.Month, tmpDate, currentproj.timeStamp) = 0 Then
                        ' in diesem Fall wurde ein Planungs-Stand im gesuchten Monat gefunden ...
                        Try

                            milestoneDate = currentproj.getMilestoneDate(milestoneName)

                        Catch ex As Exception

                            milestoneDate = awinSettings.nullDatum

                        End Try
                    Else
                        ' in diesem Fall wurde kein Planungs-Stand im gesuchten Monat gefunden ...
                        If i > 0 Then
                            ' jetzt wird gekennzeichnet, dass einfach die Bewertung des Vormonats übernommen wurde: + 12 Std
                            ' aber nur, wenn der Wert von vorher nicht schon die Kennung hatte ....
                            If DateDiff(DateInterval.Hour, tmpValues(i - 1).Date, tmpValues(i - 1)) < 7 Then
                                milestoneDate = tmpValues(i - 1).AddHours(12)
                            End If

                        Else
                            milestoneDate = awinSettings.nullDatum
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
            First = _liste.First.Value
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
            Last = _liste.Last.Value
            _currentIndex = _liste.Count - 1

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
                Throw New ArgumentException("index liegt ausserhalb der zulässigen Grenzen")
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
            If suchDatum < _liste.First.Value.timeStamp Then
                Throw New ArgumentException("es gibt keinen Projekt-Stand vor diesem Datum")
            ElseIf suchDatum > _liste.Last.Value.timeStamp Then
                _currentIndex = _liste.Count - 1
                ElementAtorBefore = _liste.Last.Value
            Else
                ' das Suchdatum liegt zwischen erstem und letztem Element

                Dim found As Boolean = False
                Dim lg As Integer = 0, rg As Integer = _liste.Count - 1
                Dim suchindex As Integer = (rg - lg) / 2

                Do While suchindex > lg And suchindex < rg And Not found

                    If _liste.ElementAt(suchindex).Key = suchDatum Then
                        found = True
                    ElseIf _liste.ElementAt(suchindex).Key > suchDatum Then
                        ' links suchen 
                        rg = suchindex
                        suchindex = lg + (rg - lg) / 2
                    Else
                        ' rechts suchen
                        lg = suchindex
                        suchindex = lg + (rg - lg) / 2

                    End If

                Loop

                If found Then
                    _currentIndex = suchindex
                    ElementAtorBefore = _liste.ElementAt(suchindex).Value
                Else
                    _currentIndex = lg
                    ElementAtorBefore = _liste.ElementAt(lg).Value
                End If


            End If



        End Get
    End Property

    ''' <summary>
    ''' fügt ein Element der Historie hinzu 
    ''' </summary>
    ''' <param name="datum"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal datum As Date, ByVal value As clsProjekt)

        Try
            If _liste.First.Value.name <> value.name Then
                Throw New ArgumentException _
                    ("Projekte mit unterschiedlichen Namen können nicht in einer Projekt-Historie sein")
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

    End Sub

End Class
