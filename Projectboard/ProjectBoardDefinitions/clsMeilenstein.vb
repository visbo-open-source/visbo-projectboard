Public Class clsMeilenstein

    Private bewertungen As SortedList(Of String, clsBewertung)
    Private _Parent As clsPhase
    Private _name As String

    Public ReadOnly Property Parent() As clsPhase
        Get
            Parent = _parent
        End Get
    End Property

    ''' <summary>
    ''' gibt den Original Namen eines Meilensteins zurück 
    ''' wenn der leer ist, dann wird der Meilenstein Name zurück gegeben 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property originalName As String
        Get

            Dim tmpNode As clsHierarchyNode
            Dim beschriftung As String = Me.name
            tmpNode = _Parent.Parent.hierarchy.nodeItem(Me.nameID)

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
    ''' setzt bzw liest die NamensID eines Meilensteins; die NamensID setzt sich zusammen aus 
    ''' dem Kennzeichen Phase/Meilenstein 0/1, dem eigentlichen Namen des Meilensteins und der laufenden Nummer. 
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
                If value.StartsWith("1§") And tmpstr.Length >= 2 Then
                    _name = value
                Else
                    Throw New ApplicationException("unzulässige Namens-ID: " & value)
                End If

            Else
                Throw New ApplicationException("Name darf nicht leer sein ...")
            End If

        End Set
    End Property

    Public Property verantwortlich As String

    Public ReadOnly Property bewertungsListe() As SortedList(Of String, clsBewertung)

        Get
            bewertungsListe = bewertungen
        End Get
    End Property



    Public ReadOnly Property getDate As Date

        Get

            Dim projektStartDate As Date
            ' das Folgende ist notwendig, um auch im Fall einer Projektvorlage ein Ergebnis zu bekommen 
            Try
                projektStartDate = Me.Parent.Parent.startDate
            Catch ex As Exception
                projektStartDate = StartofCalendar
            End Try


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

    Public Sub CopyToWithoutBewertung(ByRef newResult As clsMeilenstein)


        With newResult

            .nameID = Me.nameID
            .verantwortlich = Me.verantwortlich
            .offset = Me.offset

        End With

    End Sub


    Public Sub CopyTo(ByRef newResult As clsMeilenstein, Optional nameID As String = "")
        Dim i As Integer
        Dim newb As New clsBewertung

        With newResult

            If nameID = "" Then
                .nameID = Me.nameID
            Else
                .nameID = nameID
            End If
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
