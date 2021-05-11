
Imports Microsoft.Office.Interop.Excel
Public Class clsMeilenstein

    ' tk Änderung 4.12.17 , 
    ' - es wurde ein Attribut für percentDone aufgenommen
    ' - damit kann unterschieden werden, ob ein Meilenstein, dessen Datum in der Vergangenheit liegt, auch tatsächlich abgeschlossen wurde

    Private _percentDone As Double

    Private _nameID As String
    Private _parentPhase As clsPhase

    Private _shortName As String
    Private _originalName As String
    Private _appearance As String
    Private _color As Integer

    ' die Dokumenten Url für den Meilenstein
    Private _docURL As String

    ' die Applikations-ID mit der die Dok-Url geöffnet werden kann / soll
    Private _docUrlAppID As String

    Private _verantwortlich As String

    ' das Datum eines Meilensteines errechnet sich aus dem Phasen-Start und dem Offset ..
    Private _offset As Long


    Private _deliverables As List(Of String)
    Private _bewertungen As SortedList(Of String, clsBewertung)

    Private _invoice As KeyValuePair(Of Double, Integer)
    Private _penalty As KeyValuePair(Of Date, Double)

    ''' <summary>
    ''' liest / schreibt den Betrag, der beim Erreichen dieses Meilensteins als Rechnung gestellt werden kann 
    ''' key: Summe in T€
    ''' Value: Terms of payment
    ''' Vorsicht: kann Nothing sein. 
    ''' </summary>
    ''' <returns></returns>
    Public Property invoice As KeyValuePair(Of Double, Integer)
        Get
            invoice = _invoice
        End Get
        Set(value As KeyValuePair(Of Double, Integer))
            If Not IsNothing(value) Then
                If value.Key >= 0 And value.Value >= 0 Then
                    _invoice = value
                End If
            Else
                _invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
            End If
        End Set
    End Property

    Public Property penalty As KeyValuePair(Of Date, Double)
        Get
            penalty = _penalty
        End Get
        Set(value As KeyValuePair(Of Date, Double))
            If Not IsNothing(value) Then
                If value.Key = Date.MinValue Then
                    _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, value.Value)
                Else
                    _penalty = value
                End If

            Else
                _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0.0)
            End If
        End Set
    End Property

    ''' <summary>
    ''' es wird eine PercentDone Regelung eingeführt , mit der beurteilt werden kann, wie wit die Ergebnisse bereits sind  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property percentDone As Double
        Get
            percentDone = _percentDone
        End Get
        Set(value As Double)
            If value >= 0 Then
                If value <= 1.0 Then
                    _percentDone = value
                Else
                    _percentDone = value / 100  ' muss erst noch normiert werden, kann keine größeren Werte als 1 annehmen 
                End If

            Else
                Throw New ArgumentException("percent Done Value must not be negativ ...")
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest schreibt den String, der eine Dokumenten URL darstellt, wo Dokumente abgelegt sind, die zum Meilenstein gehören 
    ''' </summary>
    ''' <returns></returns>
    Public Property DocURL() As String
        Get
            DocURL = _docURL
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _docURL = value
            Else
                _docURL = ""
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest schreibt den String, der die ID der Appliaktion darstellt, mit der auf die Dokumenten Url zugegriffen werden kann 
    ''' </summary>
    ''' <returns></returns>
    Public Property DocUrlAppID() As String
        Get
            DocUrlAppID = _docUrlAppID
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _docUrlAppID = value
            Else
                _docUrlAppID = ""
            End If
        End Set
    End Property

    ''' <summary>
    ''' prüft zwei Meilensteine auf Identität 
    ''' </summary>
    ''' <param name="vglMS"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglMS As clsMeilenstein) As Boolean
        Get


            Dim stillOK As Boolean = False
            Dim ix As Integer = 1

            Try
                ' prüfen auf allgemeine Attribute ... 

                If Me.name = vglMS.name And
                    Me.shortName = vglMS.shortName And
                    Me.originalName = vglMS.originalName And
                    Me.appearanceName = vglMS.appearanceName And
                    Me.verantwortlich = vglMS.verantwortlich And
                    Me.offset = vglMS.offset And
                    Me.countDeliverables = vglMS.countDeliverables And
                    Me.bewertungsCount = vglMS.bewertungsCount And
                    Me.DocURL = vglMS.DocURL And
                    Me.DocUrlAppID = vglMS.DocUrlAppID And
                    Me.percentDone = vglMS.percentDone Then
                    stillOK = True


                    ' prüfen auf Deliverables ... 
                    Dim MeDelis As String = Me.getAllDeliverables("#")
                    Dim vglDelis As String = vglMS.getAllDeliverables("#")

                    If MeDelis = vglDelis Then
                        ' prüfen auf Bewertungen ... 
                        ix = 1
                        Do While stillOK And ix <= Me.bewertungsCount
                            Dim MeBewertung As clsBewertung = Me.getBewertung(ix)
                            Dim vglBewertung As clsBewertung = vglMS.getBewertung(ix)
                            If MeBewertung.isIdenticalTo(vglBewertung) Then
                                ix = ix + 1
                            Else
                                stillOK = False
                            End If
                        Loop

                    End If



                End If

                ' jetzt die Invoices und Penalties abfragen 
                If stillOK Then
                    stillOK = Me.invoice.Key = vglMS.invoice.Key And
                        Me.invoice.Value = vglMS.invoice.Value And
                        Me.penalty.Key = vglMS.penalty.Key And
                        Me.penalty.Value = vglMS.penalty.Value
                End If


            Catch ex As Exception
                stillOK = False
            End Try



            isIdenticalTo = stillOK

        End Get
    End Property
    ' Farbe, Form und Abkürzung eines Meilensteins wird über den categorizedName bzw. die missingmilestonedefinitions abgebildet 
    ' oder aber über den die folgenden Parameter 

    ''' <summary>
    ''' gibt die Anzahl Deliverables für diesen Meilenstein zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property countDeliverables As Integer
        Get
            countDeliverables = _deliverables.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt an , ob das Deliverable existiert ...
    ''' </summary>
    ''' <param name="item"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsDeliverable(ByVal item As String) As Boolean
        Get
            containsDeliverable = _deliverables.Contains(item)
        End Get
    End Property

    ''' <summary>
    ''' fügt das Deliverable Item der Liste hinzu; 
    ''' wenn das Item bereits in der Liste vorhanden ist, passiert nichts 
    ''' </summary>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Public Sub addDeliverable(ByVal item As String)

        If Not _deliverables.Contains(item) Then
            _deliverables.Add(item)
        End If

    End Sub

    ''' <summary>
    ''' löscht alle Deliverables des Meilensteines 
    ''' </summary>
    Public Sub clearDeliverables()
        _deliverables.Clear()
    End Sub
    ''' <summary>
    ''' gibt das Element an der bezeichneten Stelle zurück
    ''' index kann Werte zwischen 1 .. count annehmen 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDeliverable(ByVal index As Integer) As String
        Get
            Dim tmpValue As String = ""
            If index >= 1 And index <= _deliverables.Count Then
                tmpValue = _deliverables.Item(index - 1)
            End If
            getDeliverable = tmpValue
        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der Deliverables eines Meilensteins als einen String zurück; 
    ''' die einzelnen Deliverables sind by default durch einen vblf getrennt
    ''' oder getrennt durch das übergebene trennzeichen  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllDeliverables(ByVal trennzeichen As String) As String
        Get
            Dim tmpDeliverables As String = ""
            Dim tmp_i As Integer = 1

            For i As Integer = 1 To _deliverables.Count
                ' ur:07.02.2020 nur nicht leere Deliverables sind relevant
                If _deliverables.Item(i - 1) <> "" Then
                    If tmp_i = 1 Then
                        tmpDeliverables = _deliverables.Item(i - 1)
                        tmp_i = tmp_i + 1
                    Else
                        tmpDeliverables = tmpDeliverables & trennzeichen &
                            _deliverables.Item(i - 1)
                        tmp_i = tmp_i + 1
                    End If
                End If

            Next

            getAllDeliverables = tmpDeliverables


        End Get
    End Property


    ''' <summary>
    ''' liest / setzt die individuelle appearance für diesen Meilenstein 
    ''' normalerweise wird die Appearance aber über die MilestoneDefinitions oder missingMilestoneDefinitions definiert 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property appearanceName As String
        Get

            '' tk/ute. 29.11.20
            'If MilestoneDefinitions.Contains(Me.name) Then
            '    _appearance = MilestoneDefinitions.getAppearance(Me.name)
            'End If
            'If _appearance = "" Then
            '    _appearance = awinSettings.defaultMilestoneClass
            'End If
            appearanceName = _appearance

        End Get
        Set(value As String)
            If appearanceDefinitions.liste.ContainsKey(value) Then
                _appearance = value
            Else
                _appearance = awinSettings.defaultMilestoneClass
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest/schreibt den Ampel-Status, das ist die 1. Bewertung
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ampelStatus As Integer
        Get
            If Me.bewertungsCount >= 1 Then
                ampelStatus = Me.getBewertung(1).colorIndex
            Else
                ampelStatus = 0
            End If
        End Get

        Set(value As Integer)
            If IsNothing(value) Then
                value = 0
            ElseIf value < 0 Or value > 3 Then
                value = 0
            End If

            If Me.bewertungsCount >= 1 Then
                Me.getBewertung(1).colorIndex = value
            Else

                Dim tmpB As New clsBewertung
                With tmpB
                    .description = ""
                    .colorIndex = value
                End With

                Me.addBewertung(tmpB)

            End If
        End Set

    End Property

    ''' <summary>
    ''' liest/schreibt die Ampel-Erläuterung, das ist die 1. Bewertung
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ampelErlaeuterung As String
        Get
            If Me.bewertungsCount >= 1 Then
                ampelErlaeuterung = Me.getBewertung(1).description
            Else
                ampelErlaeuterung = ""
            End If
        End Get
        Set(value As String)
            If IsNothing(value) Then
                value = ""
            End If

            If Me.bewertungsCount >= 1 Then
                Me.getBewertung(1).description = value
            Else
                Dim tmpB As New clsBewertung
                With tmpB
                    .description = value
                    .colorIndex = 0
                End With

                Me.addBewertung(tmpB)

            End If

        End Set

    End Property

    ''' <summary>
    ''' gibt die individuelle Farbe zurück, also die Einstellung, die verwendet wird 
    ''' wenn es sich nicht um einen categorized namen handelt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property individualColor As Integer
        Get
            individualColor = Me._color
        End Get
    End Property


    ''' <summary>
    ''' gibt die Farbe eines Meilensteins zurück; wenn er in der Liste der bekannten Meilensteine ist, 
    ''' dann die Farbe der Darstellungsklasse, sonst die AlternativeFare, die ggf beim auslesen aus MS Project ermittelt wird
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property farbe As Integer
        Get
            Try

                Dim msName As String = elemNameOfElemID(_nameID)

                If Not IsNothing(appearanceDefinitions.getMileStoneAppearance(name, appearanceName)) Then
                    'ur:190725
                    'farbe = Me.getShape.Fill.ForeColor.RGB
                    farbe = appearanceDefinitions.getMileStoneAppearance(name, appearanceName).FGcolor
                Else
                    farbe = _color
                End If

            Catch ex As Exception
                farbe = _color
            End Try

        End Get
        Set(value As Integer)
            If value >= RGB(0, 0, 0) And value <= RGB(255, 255, 255) Then
                _color = value
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
            Parent = _parentPhase
        End Get
    End Property

    ''' <summary>
    ''' liest/schreibt den Original Name
    ''' gibt den Original Namen eines Meilensteins zurück 
    ''' wenn der leer ist, dann wird der Name zurück gegeben 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property originalName As String
        Get

            If _originalName = "" Then
                originalName = Me.name
            Else
                originalName = _originalName
            End If

        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _originalName = value
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' gibt die Abkürzung des Meilensteins zurück 
    ''' entweder als Abkürzung der phaseDefinitions, als Abkürzung der missingphaseDefinitions oder der leere String
    ''' Später: alternativeAbbrev
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property shortName As String
        Get
            Dim abbrev As String = ""
            Dim tmpName As String = Me.name

            If MilestoneDefinitions.Contains(tmpName) Then
                abbrev = MilestoneDefinitions.getAbbrev(tmpName)
            ElseIf missingMilestoneDefinitions.Contains(tmpName) Then
                abbrev = missingMilestoneDefinitions.getAbbrev(tmpName)
            Else
                abbrev = _shortName
            End If

            shortName = abbrev

        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _shortName = value
                End If
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
            name = elemNameOfElemID(_nameID)
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
            nameID = _nameID
        End Get
        Set(value As String)
            Dim tmpstr() As String
            tmpstr = value.Split(New Char() {CChar("§")}, 3)
            If Len(value) > 0 Then
                If value.StartsWith("1§") And tmpstr.Length >= 2 Then
                    _nameID = value
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
        Get
            verantwortlich = _verantwortlich
        End Get
        Set(value As String)
            _verantwortlich = value
        End Set
    End Property

    ''' <summary>
    ''' gibt die Bewertungsliste zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property bewertungsListe() As SortedList(Of String, clsBewertung)

        Get
            bewertungsListe = _bewertungen
        End Get
    End Property

    Public ReadOnly Property getPenaltyValue As Double
        Get
            getPenaltyValue = _penalty.Value
        End Get
    End Property
    Public ReadOnly Property getPenaltyDate As Date
        Get
            getPenaltyDate = _penalty.Key
        End Get
    End Property

    Public ReadOnly Property getPaymentValue As Double
        Get
            getPaymentValue = _invoice.Key
        End Get
    End Property
    ''' <summary>
    ''' gibt das Datum des vorauss Geldeingangs wieder
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getPaymentDate As Date
        Get
            getPaymentDate = getDate.AddDays(_invoice.Value)
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
                projektStartDate = Me.Parent.parentProject.startDate
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

            ' Änderung tk, 20.6.18 .startdate.Date, um zu normieren  
            Dim projektStartDate As Date = Me.Parent.parentProject.startDate.Date
            Dim phasenOffset As Integer = Me.Parent.startOffsetinDays

            If DateDiff(DateInterval.Day, projektStartDate, value.Date) < 0 Then
                Throw New Exception("ungültiges Datum für Meilenstein " & value.ToShortDateString)

            Else
                Try
                    ' Änderung tk, 20.6.18 value.date , um zu normieren ...
                    _offset = DateDiff(DateInterval.Day, projektStartDate.AddDays(phasenOffset).Date, value.Date)
                Catch ex As Exception
                    Throw New Exception("ungültiges Datum für Meilenstein " & value.ToShortDateString & vbLf &
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
            _bewertungen.Clear()
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
        Get
            offset = _offset
        End Get
        Set(value As Long)
            _offset = value
        End Set
    End Property


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property bewertungsCount As Integer
        Get
            bewertungsCount = _bewertungen.Count

        End Get
    End Property

    ''Public Sub CopyToWithoutBewertung(ByRef newResult As clsMeilenstein)


    ''    With newResult

    ''        .nameID = Me.nameID
    ''        .verantwortlich = Me.verantwortlich
    ''        .offset = Me.offset
    ''        .farbe = Me.farbe

    ''    End With

    ''End Sub


    Public Sub copyTo(ByRef newResult As clsMeilenstein, Optional optNameID As String = "")
        Dim i As Integer


        With newResult

            .offset = Me._offset


            If optNameID = "" Then
                .nameID = Me.nameID
            Else
                .nameID = optNameID
            End If


            .shortName = Me._shortName
            .originalName = Me._originalName
            .appearanceName = Me._appearance
            .farbe = Me._color
            .verantwortlich = Me._verantwortlich
            .percentDone = Me._percentDone

            ' tk 2.6.20
            .invoice = _invoice
            .penalty = _penalty


            For i = 1 To Me._bewertungen.Count
                Dim newb As New clsBewertung
                Me.getBewertung(i).copyto(newb)
                Try
                    .addBewertung(newb)
                Catch ex As Exception

                End Try

            Next

            ' jetzt noch die Deliverables kopieren ... 
            For i = 1 To Me.countDeliverables
                Dim deli As String = Me.getDeliverable(i)
                .addDeliverable(deli)
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
            If _bewertungen.ContainsKey(key) Then
                _bewertungen.Remove(key)
            End If
            _bewertungen.Add(key, b)
        Catch ex As Exception
            Throw New ArgumentException("Bewertung wurde bereites vergeben ..")
        End Try

    End Sub

    Public Sub removeBewertung(ByVal key As String)

        Try
            _bewertungen.Remove(key)
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

    End Sub

    Public ReadOnly Property getBewertung(ByVal index As Integer) As clsBewertung

        Get

            If index > _bewertungen.Count Then
                'getBewertung = Nothing
                getBewertung = New clsBewertung
            Else
                getBewertung = _bewertungen.ElementAt(index - 1).Value
            End If

        End Get

    End Property


    Sub New(ByRef parent As clsPhase)

        _nameID = ""
        _parentPhase = parent

        ' Vorbesetzen der Dokumenten-URL und App-ID , mit der die Dokumente bearbeitet werden können 
        _docURL = ""
        _docUrlAppID = ""

        _percentDone = 0.0
        _bewertungen = New SortedList(Of String, clsBewertung)
        _deliverables = New List(Of String)

        _shortName = ""
        _originalName = ""
        _appearance = awinSettings.defaultMilestoneClass

        Try
            _color = XlRgbColor.rgbAquamarine
            'If appearanceDefinitions.ContainsKey(_appearance) Then
            '    If Not IsNothing(appearanceDefinitions.Item(_appearance).form) Then
            '        _color = appearanceDefinitions.Item(_appearance).form.Fill.ForeColor.RGB
            '    End If
            'End If
        Catch ex As Exception

        End Try

        _verantwortlich = ""

        _invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
        _penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)

        offset = 0
        


    End Sub

End Class
