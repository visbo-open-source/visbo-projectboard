Public Class clsResult

    Private bewertungen As SortedList(Of String, clsBewertung)
    Private _Parent As clsPhase

    Public ReadOnly Property Parent() As clsPhase
        Get
            Parent = _parent
        End Get
    End Property


    Public Property name As String

    Public Property verantwortlich As String

    Public ReadOnly Property bewertungsListe() As SortedList(Of String, clsBewertung)

        Get
            bewertungsListe = bewertungen
        End Get
    End Property



    Public ReadOnly Property getDate As Date

        Get

            Dim projektStartDate As Date = Me.Parent.Parent.startDate
            Dim phasenOffset As Integer = Me.Parent.startOffsetinDays

            getDate = projektStartDate.AddDays(phasenOffset + _offset)

        End Get

    End Property

    Public WriteOnly Property setDate As Date

        Set(value As Date)

            Dim projektStartDate As Date = Me.Parent.Parent.startDate
            Dim phasenOffset As Integer = Me.Parent.startOffsetinDays

            If DateDiff(DateInterval.Day, projektStartDate, value) < 0 Then
                Throw New Exception("ungültiges Datum für Meilenstein " & value.ToShortDateString)

            Else
                Try
                    _offset = DateDiff(DateInterval.Day, projektStartDate.AddDays(phasenOffset), value)
                Catch ex As Exception
                    Throw New Exception("ungültiges Datum für Meilenstein " & value.ToShortDateString & vbLf & _
                                        ex.Message)
                End Try

            End If

        End Set

    End Property

    'Public Sub setDate(ByVal parentStartDate As Date, ByVal resultDate As Date)

    '    Try
    '        _offset = DateDiff(DateInterval.Day, parentStartDate, resultDate)
    '    Catch ex As Exception
    '        _offset = 0
    '    End Try


    'End Sub

    Public Sub clearBewertungen()

        Try
            bewertungen.Clear()
        Catch ex As Exception

        End Try

    End Sub


    Public Property offset As Long

    'Friend Property fileLink As Uri

    Public ReadOnly Property bewertungsCount As Integer
        Get
            bewertungsCount = bewertungen.Count

        End Get
    End Property

    Public Sub CopyToWithoutBewertung(ByRef newResult As clsResult)

        
        With newResult

            .name = Me.name
            .verantwortlich = Me.verantwortlich
            .offset = Me.offset

        End With

    End Sub


    Public Sub CopyTo(ByRef newResult As clsResult)
        Dim i As Integer
        Dim newb As New clsBewertung

        With newResult

            .name = Me.name
            .verantwortlich = Me.verantwortlich
            .offset = Me.offset

            For i = 1 To Me.bewertungen.Count
                Me.getBewertung(i).copyto(newb)
                Try
                    .addBewertung(newb)
                Catch ex As Exception

                End Try

            Next

        End With

    End Sub


    Public Sub addBewertung(ByVal b As clsBewertung)
        Dim key As String

        If Not b.bewerterName Is Nothing Then
            key = b.bewerterName & "#" & b.datum.ToString("MMM yy")
        Else
            key = "#" & b.datum.ToString("MMM yy")
        End If

        Try
            bewertungen.Add(key, b)
        Catch ex As Exception
            Throw New ArgumentException("Bewertung wurde bereites vergeben ..")
        End Try

    End Sub

    Public Sub removeBewertung(ByVal key As String)

        Try
            bewertungen.Remove(key)
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

    End Sub

    Public ReadOnly Property getBewertung(ByVal index As Integer) As clsBewertung

        Get

            If index > bewertungen.Count Then
                'getBewertung = Nothing
                getBewertung = New clsBewertung
            Else
                getBewertung = bewertungen.ElementAt(index - 1).Value
            End If

            'Try
            '    getBewertung = bewertungen.ElementAt(index - 1).Value
            'Catch ex As Exception
            '    getBewertung = Nothing
            '    Throw New ArgumentException(ex.Message)
            'End Try

        End Get

    End Property

    Public ReadOnly Property getBewertung(ByVal key As String) As clsBewertung

        Get


            Try
                getBewertung = bewertungen.Item(key)
            Catch ex As Exception
                getBewertung = Nothing
                Throw New ArgumentException(ex.Message)
            End Try

        End Get

    End Property

    'Sub New()

    '    bewertungen = New SortedList(Of String, clsBewertung)
    '    _offset = 0

    'End Sub

    Sub New(ByRef parent As clsPhase)

        bewertungen = New SortedList(Of String, clsBewertung)
        _offset = 0
        _Parent = parent

    End Sub

End Class
