Public Class clsPPTTimeMachine

    Private _timeStamps As SortedList(Of Date, Boolean)
    ' tk 28.10.18 wird nicht mehr gebraucht ... 
    'Private _timeStampsIndex As Integer
    Private _anzahlShapesOnSlide As Integer

    Public Property timeStamps As SortedList(Of Date, Boolean)
        Get
            timeStamps = _timeStamps
        End Get
        Set(value As SortedList(Of Date, Boolean))
            If Not IsNothing(value) Then
                _timeStamps = value
            Else
                _timeStamps = Nothing
            End If
        End Set
    End Property

    'Public Property timeStampsIndex As Integer
    '    Get
    '        timeStampsIndex = _timeStampsIndex
    '    End Get
    '    Set(value As Integer)
    '        If Not IsNothing(value) Then
    '            If IsNumeric(value) Then
    '                _timeStampsIndex = value
    '            Else
    '                _timeStampsIndex = -1
    '            End If
    '        Else
    '            _timeStampsIndex = -1
    '        End If
    '    End Set
    'End Property

    ''' <summary>
    ''' fügt der leeren oder bereits existierenden timeStampliste die neuen Werte hinzu
    ''' wenn leer, wird erstmal die Gesamte Liste aufgenommen , danach nur noch evtl Min/Max Ausreisser  
    ''' </summary>
    ''' <param name="tsCollection"></param>
    Public Sub addNewList(ByVal tsCollection As Collection)
        Dim wasEmpty As Boolean = _timeStamps.Count = 0

        For Each tmpDate As Date In tsCollection

            If wasEmpty Then

                If Not _timeStamps.ContainsKey(tmpDate) Then
                    _timeStamps.Add(tmpDate, False)
                End If

            Else
                ' jetzt wird erstmal lediglich geprüft, ob das Datum kleiner als der kleinste Werte oder größer als der größte Werte ist 
                ' wenn eines davon zutrifft, wird es aufgenommen. 
                ' Damit wird sichergestellt, dass die TimeStamps Liste zwar nicht alle Timestamps enthält, aber die Min, Max Werte korrekt gesetzt sind 

                If DateDiff(DateInterval.Minute, tmpDate, _timeStamps.First) > 0 Then
                    _timeStamps.Add(tmpDate, False)
                ElseIf DateDiff(DateInterval.Minute, _timeStamps.Last, tmpDate) > 0 Then
                    _timeStamps.Add(tmpDate, False)
                End If
            End If


        Next

    End Sub

    Private Enum ptNavigationButtons
        letzter = 0
        erster = 1
        nachher = 2
        vorher = 3
        individual = 4
    End Enum


    Public Sub New()
        timeStamps = New SortedList(Of Date, Boolean)
        'timeStampsIndex = -1
    End Sub
End Class
