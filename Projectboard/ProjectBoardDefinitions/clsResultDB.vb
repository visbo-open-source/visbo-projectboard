Public Class clsResultDB

    Public bewertungen As SortedList(Of String, clsBewertungDB)


    Public name As String
    Public verantwortlich As String
    Public offset As Long
    Public alternativeColor As Long

    Public shortName As String
    Public originalName As String
    Public appearance As String

    Public deliverables As List(Of String)
    Public percentDone As Double

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


    Friend Sub CopyTo(ByRef newResult As clsMeilenstein)
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
                            .addDeliverable(tmpDeliverable)
                        Next
                    Else
                        ' evtl sind die noch in der Bewertung vergraben ... 
                        If Me.bewertungsCount > 0 Then
                            If Not IsNothing(Me.getBewertung(1).deliverables) Then
                                Dim allDeliverables As String = Me.getBewertung(1).deliverables

                                If allDeliverables.Trim.Length > 0 Then
                                    Dim tmpstr() As String = allDeliverables.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)
                                    For i = 1 To tmpstr.Length
                                        .addDeliverable(tmpstr(i - 1))
                                    Next
                                End If

                            End If
                        End If
                    End If
                End If

                For i = 1 To Me.bewertungsCount

                    Dim newb As New clsBewertung
                    Try
                        Me.getBewertung(i).CopyTo(newb)
                        .addBewertung(newb)
                    Catch ex1 As Exception

                    End Try

                Next

            End With

        Catch ex As Exception

        End Try



    End Sub

    Friend Sub CopyFrom(ByVal newResult As clsMeilenstein)
        Dim i As Integer


        With newResult

            Me.name = .nameID
            Me.verantwortlich = .verantwortlich
            Me.offset = .offset

            Me.shortName = .shortName
            Me.originalName = .originalName
            Me.appearance = .appearance

            Me.alternativeColor = .individualColor

            Me.percentDone = .percentDone
           
            For i = 1 To .countDeliverables
                Dim tmpDeliverable As String = .getDeliverable(i)
                Me.deliverables.Add(tmpDeliverable)
            Next

            Try
                For i = 1 To .bewertungsCount
                    Dim newb As New clsBewertungDB
                    newb.Copyfrom(.getBewertung(i))
                    Me.addBewertung(newb)
                Next
            Catch ex As Exception

            End Try


        End With

    End Sub


    Friend Sub addBewertung(ByVal b As clsBewertungDB)
        Dim key As String

        If Not b.bewerterName Is Nothing Then
            key = b.bewerterName.Trim & "#" & b.datum.ToString("MMM yy")
        Else
            key = "#" & b.datum.ToString("MMM yy")
        End If

        Try
            bewertungen.Add(key, b)
        Catch ex As Exception

            Throw New ArgumentException("Bewertung wurde bereits vergeben ..")

        End Try

    End Sub

    Friend Sub removeBewertung(ByVal key As String)

        Try
            bewertungen.Remove(key)
        Catch ex As Exception

            Throw New ArgumentException(ex.Message)

        End Try

    End Sub

    Friend ReadOnly Property getBewertung(ByVal index As Integer) As clsBewertungDB

        Get

            Try
                getBewertung = bewertungen.ElementAt(index - 1).Value
            Catch ex As Exception
                getBewertung = Nothing
                Throw New ArgumentException(ex.Message)
            End Try


        End Get

    End Property

    Sub New()

        percentDone = 0.0
        bewertungen = New SortedList(Of String, clsBewertungDB)
        deliverables = New List(Of String)

    End Sub


End Class
