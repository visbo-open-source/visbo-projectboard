Public Class clsMeilenstein

    Private bewertungen As SortedList(Of String, clsBewertung)
    Private _Parent As clsPhase
    Private _name As String

    ' Erweiterung tk 18.2.16
    ' das wird verwendet . um eine Farbe Meilensteins, der nicht zur Liste der bekannten gehört 
    ' aufzunehmen 
    Private _alternativeColor As Long

    ''' <summary>
    ''' gibt die Farbe eines Meilensteins zurück; wenn er in der Liste der bekannten Meilensteine ist, 
    ''' dann die Farbe der Darstellungsklasse, sonst die AlternativeFare, die ggf beim auslesen aus MS Project ermittelt wird
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property farbe As Long
        Get
            Dim msName As String = elemNameOfElemID(_name)
            If MilestoneDefinitions.Contains(msName) Then
                farbe = CLng(MilestoneDefinitions.getShape(msName).Fill.ForeColor.RGB)
            Else
                farbe = _alternativeColor
            End If
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
    ''' gibt die Eltern-Phase zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Parent() As clsPhase
        Get
            Parent = _Parent
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

    ''' <summary>
    ''' liest/schreibt wer verantwortlich ist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property verantwortlich As String

    ''' <summary>
    ''' gibt die Bewertungsliste zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property bewertungsListe() As SortedList(Of String, clsBewertung)

        Get
            bewertungsListe = bewertungen
        End Get
    End Property



    ''' <summary>
    ''' liest das Datum des Meilensteins
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' setzt das Datum des Meilensteins, d.h intern wird die Variable _offset gesetzt 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
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


    ''' <summary>
    ''' löscht die Bewertungen des Meilensteins
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearBewertungen()

        Try
            bewertungen.Clear()
        Catch ex As Exception

        End Try

    End Sub


    ''' <summary>
    ''' setzt / liest den Offset, das heisst den Abstand in Tagen vom Phasen-Start bis zum Meilenstein 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property offset As Long


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
            .setFarbe = Me.farbe

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
            .setFarbe = Me.farbe

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

        Dim defaultName As String = "Meilenstein Default"
        bewertungen = New SortedList(Of String, clsBewertung)
        _offset = 0
        _Parent = parent
        _alternativeColor = awinSettings.AmpelNichtBewertet
        
    End Sub

End Class
