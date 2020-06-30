Imports ProjectBoardDefinitions
''' <summary>
''' Klassendefinition für einen Meilenstein Zugriff über ReST
''' </summary>
Public Class clsResultWeb

    Public bewertungen As List(Of clsBewertungWeb)

    Public name As String
    Public verantwortlich As String
    Public offset As Long
    Public alternativeColor As Long

    Public shortName As String
    Public originalName As String
    Public appearance As String

    Public docURL As String
    Public docUrlAppID As String

    Public deliverables As List(Of String)
    Public percentDone As Double

    ' tk 2.6.2020
    Public invoice As KeyValuePair(Of Double, Integer)
    Public penalty As KeyValuePair(Of Date, Double)

    'Friend Property fileLink As Uri

    Friend ReadOnly Property bewertungsCount As Integer

        Get

            Try
                bewertungsCount = bewertungen.Count
            Catch ex As Exception
                bewertungsCount = 0
            End Try

        End Get
    End Property


    Public Sub CopyTo(ByRef newResult As clsMeilenstein)
        Dim i As Integer

        Try
            With newResult

                ' Änderung tk 20.4.2015
                ' damit alte Datenbank Einträge ohne Hierarchie auch noch gelesen werden können ..
                If Not istElemID(Me.name) Then
                    .nameID = calcHryElemKey(Me.name, True)
                Else
                    .nameID = Me.name
                End If

                ' jetzt kommen die Doumenten Folder und AppIDs
                If Not IsNothing(Me.docURL) Then
                    .DocURL = Me.docURL
                End If

                If Not IsNothing(Me.docUrlAppID) Then
                    .DocUrlAppID = Me.docUrlAppID
                End If

                .verantwortlich = Me.verantwortlich
                .offset = Me.offset

                If Not IsNothing(Me.shortName) Then
                    .shortName = Me.shortName
                End If

                If Not IsNothing(Me.originalName) Then
                    .originalName = Me.originalName
                End If

                If Not IsNothing(Me.appearance) Then
                    .appearance = Me.appearance
                End If

                If Not IsNothing(Me.percentDone) Then
                    .percentDone = Me.percentDone
                End If

                If Not IsNothing(Me.invoice) Then
                    .invoice = Me.invoice
                End If

                If Not IsNothing(Me.penalty) Then
                    .penalty = Me.penalty
                End If
                Try
                    If Not IsNothing(Me.alternativeColor) Then
                        .farbe = CInt(Me.alternativeColor)
                    Else
                        .farbe = CInt(awinSettings.AmpelNichtBewertet)
                    End If
                Catch ex As Exception

                End Try

                If Not IsNothing(Me.deliverables) Then
                    If Me.deliverables.Count > 0 Then
                        For i = 1 To Me.deliverables.Count
                            Dim tmpDeliverable As String = Me.deliverables.Item(i - 1)
                            'ur:07.02.2020: nur nicht leere Deliverables mitnehmen
                            '          .addDeliverable(tmpDeliverable)
                            If tmpDeliverable <> "" Then
                                .addDeliverable(tmpDeliverable)
                            End If
                        Next
                    Else
                        ' evtl sind die noch in der Bewertung vergraben ... 
                        If Me.bewertungsCount > 0 Then
                            If Not IsNothing(Me.getBewertung(1).bewertung.deliverables) Then
                                Dim allDeliverables As String = Me.getBewertung(1).bewertung.deliverables

                                If allDeliverables.Trim.Length > 0 Then
                                    Dim tmpstr() As String = allDeliverables.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)
                                    For i = 1 To tmpstr.Length
                                        ' ur:07.02.2020: nur nicht leere Deliverables mitnehmen
                                        If tmpstr(i - 1) <> "" Then
                                            .addDeliverable(tmpstr(i - 1))
                                        End If
                                    Next
                                End If

                            End If
                        End If
                    End If
                End If


                For Each wBew As clsBewertungWeb In Me.bewertungen
                    Dim newb As New clsBewertung

                    Dim nbewkey As String = wBew.key
                    wBew.bewertung.CopyTo(newb)
                    Try
                        .addBewertung(newb)
                        '.bewertungsListe.Add(nbewkey, newb)
                    Catch ex As Exception

                    End Try
                Next

            End With

        Catch ex As Exception

        End Try



    End Sub

    Public Sub CopyFrom(ByVal newResult As clsMeilenstein)
        Dim i As Integer


        With newResult

            Me.name = .nameID
            Me.verantwortlich = .verantwortlich
            Me.offset = .offset

            Me.shortName = .shortName
            Me.originalName = .originalName
            Me.appearance = .appearance

            ' Dokumenten-Url und Applikation
            Me.docURL = .DocURL
            Me.docUrlAppID = .DocUrlAppID

            Me.alternativeColor = .individualColor

            Me.percentDone = .percentDone

            'Me.invoice = .invoice
            Me.invoice = New KeyValuePair(Of Double, Integer)(10.5, 30)
            'Me.penalty = .penalty
            Me.penalty = New KeyValuePair(Of Date, Double)(Date.MinValue, 10.5)

            For i = 1 To .countDeliverables
                Dim tmpDeliverable As String = .getDeliverable(i)
                'ur:07.02.2020: nur nicht leere Deliverables mitnehmen
                '   Me.deliverables.Add(tmpDeliverable)
                If tmpDeliverable <> "" Then
                    Me.deliverables.Add(tmpDeliverable)
                End If
            Next

            Try    ' evtl vorhandene Bewertungen kopieren .... 

                For Each kvp As KeyValuePair(Of String, clsBewertung) In .bewertungsListe
                    Dim newb As New clsBewertungWeb
                    newb.key = kvp.Key
                    newb.bewertung.Copyfrom(kvp.Value)
                    Me.addBewertung(newb)
                Next

            Catch ex As Exception

            End Try


        End With

    End Sub


    Friend Sub addBewertung(ByVal b As clsBewertungWeb)
        Dim key As String = b.key

        If Not b.bewertung.bewerterName Is Nothing Then
            key = b.bewertung.bewerterName.Trim & "#" & b.bewertung.datum.ToString("MMM yy")
        Else
            key = "#" & b.bewertung.datum.ToString("MMM yy")
        End If

        Try
            bewertungen.Add(b)
        Catch ex As Exception

            Throw New ArgumentException("Bewertung wurde bereits vergeben ..")

        End Try

    End Sub

    Friend Sub removeBewertung(ByVal key As String)

        Try
            For Each wbew As clsBewertungWeb In bewertungen
                If wbew.key = key Then
                    bewertungen.Remove(wbew)
                    Exit For
                End If
            Next

        Catch ex As Exception

            Throw New ArgumentException(ex.Message)

        End Try

    End Sub

    Friend ReadOnly Property getBewertung(ByVal index As Integer) As clsBewertungWeb

        Get

            Try
                getBewertung = bewertungen.Item(index - 1)
            Catch ex As Exception
                getBewertung = Nothing
                Throw New ArgumentException(ex.Message)
                Call MsgBox("Fehler in .getBewertung")
            End Try


        End Get

    End Property

    Sub New()

        percentDone = 0.0
        bewertungen = New List(Of clsBewertungWeb)
        deliverables = New List(Of String)
        docURL = ""
        docUrlAppID = ""

        invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
        penalty = New KeyValuePair(Of Date, Double)(Date.MinValue, 0.0)

    End Sub


End Class
