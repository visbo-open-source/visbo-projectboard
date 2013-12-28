Public Class clsProjektHistorie

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

    Public ReadOnly Property Count As Integer
        Get
            Count = _liste.Count
        End Get
    End Property

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

    Public ReadOnly Property item(ByVal index As Integer) As clsProjekt
        Get

            If index >= 0 And index <= _liste.Count - 1 Then
                ' ein element mit diesem Index existiert ... 
                item = _liste.ElementAt(index).Value
                _currentIndex = index
            Else
                Throw New ArgumentException("index liegt ausserhalb der zulässigen Grenze: " & index)
            End If

        End Get
    End Property

    Public Sub clear()

        _liste.Clear()

    End Sub

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

    Public ReadOnly Property First As clsProjekt
        Get
            First = _liste.First.Value
            _currentIndex = 0
        End Get
    End Property

    Public ReadOnly Property Last As clsProjekt
        Get
            Last = _liste.Last.Value
            _currentIndex = _liste.Count - 1

        End Get
    End Property

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
