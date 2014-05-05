Public Class clsProjektDB

    Public name As String
    Public variantName As String
    Public Risiko As Double
    Public StrategicFit As Double
    
    Public Erloes As Double
    Public leadPerson As String
    Public tfSpalte As Integer
    Public tfZeile As Integer
    Public startDate As Date
    Public endDate As Date
    Public earliestStart As Integer
    Public earliestStartDate As Date
    Public latestStart As Integer
    Public latestStartDate As Date
    Public status As String
    Public ampelStatus As Integer
    Public ampelErlaeuterung As String
    Public farbe As Object
    Public Schrift As Integer
    Public Schriftfarbe As Object
    Public VorlagenName As String
    Public Dauer As Integer
    Public AllPhases As List(Of clsPhaseDB)
    Public Id As String
    Public timestamp As Date
    ' ergänzt am 16.11.13
    Public volumen As Double
    Public complexity As Double
    Public description As String
    Public businessUnit As String

    Public Sub copyfrom(ByVal projekt As clsProjekt)
        Dim i As Integer


        'Me.timestamp = Date.Now
        'Me.Id = 0

        With projekt
            ' damit alle Projekte die gleiche Timestamp für das Datenbank Speichern haben wird das in der 
            ' aufrufenden Sequenz erledigt Me.timestamp = Date.UtcNow
            If Not IsNothing(.timeStamp) Then
                Me.timestamp = .timeStamp.ToUniversalTime
            Else
                Me.timestamp = Date.UtcNow
            End If

            If Not IsNothing(.Id) Then
                Me.Id = .Id
            End If
            Me.name = .name
            Me.variantName = .variantName
            Me.Risiko = .Risiko
            Me.StrategicFit = .StrategicFit
            Me.Erloes = .Erloes
            Me.leadPerson = .leadPerson
            Me.tfSpalte = .tfspalte
            Me.tfZeile = .tfZeile
            Me.startDate = .startDate.ToUniversalTime
            Me.endDate = .endeDate.ToUniversalTime
            Me.earliestStartDate = .earliestStartDate.ToUniversalTime
            Me.latestStartDate = .latestStartDate.ToUniversalTime
            Me.earliestStart = .earliestStart
            Me.latestStart = .latestStart
            Me.status = .Status
            Me.ampelStatus = .ampelStatus
            Me.ampelErlaeuterung = .ampelErlaeuterung
            Me.farbe = .farbe
            Me.Schrift = .Schrift
            Me.Schriftfarbe = .Schriftfarbe
            Me.VorlagenName = .VorlagenName
            Me.Dauer = .Dauer
            ' ergänzt am 16.11.13
            Me.volumen = .volume
            Me.complexity = .complexity
            Me.description = .description
            Me.businessUnit = .businessUnit


            For i = 1 To .CountPhases
                Dim newPhase As New clsPhaseDB
                newPhase.copyFrom(.getPhase(i), .farbe)
                AllPhases.Add(newPhase)
            Next

        End With

    End Sub

    Public Sub copyto(ByRef projekt As clsProjekt)
        Dim i As Integer


        With projekt
            .timeStamp = Me.timestamp.ToLocalTime
            .Id = Me.Id
            .name = Me.name
            .variantName = Me.variantName
            .Risiko = Me.Risiko
            .StrategicFit = Me.StrategicFit
            .Erloes = Me.Erloes
            .leadPerson = Me.leadPerson
            ' es gibt kein Attribut tfspalte mehr - es ist ein Readonly Attribut, wo _Start ausgelesen wird 
            '.tfSpalte = Me.tfSpalte
            .tfZeile = Me.tfZeile
            .startDate = Me.startDate.ToLocalTime
            .earliestStartDate = Me.earliestStartDate.ToLocalTime
            .latestStartDate = Me.latestStartDate.ToLocalTime
            .earliestStart = Me.earliestStart
            .latestStart = Me.latestStart
            .Status = Me.status
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung
            .farbe = Me.farbe
            .Schrift = Me.Schrift

            .volume = Me.volumen
            .complexity = Me.complexity
            .description = Me.description
            .businessUnit = Me.businessUnit

            ' Änderung notwendig, weil mal in der Datenbank Schrift mit -10 stand
            If .Schrift < 0 Then
                .Schrift = -1 * .Schrift
            End If
            .Schriftfarbe = Me.Schriftfarbe
            .VorlagenName = Me.VorlagenName
            '.Dauer = Me.Dauer
            For i = 1 To Me.AllPhases.Count
                Dim newPhase As New clsPhase(projekt)
                AllPhases.Item(i - 1).copyto(newPhase)
                .AddPhase(newPhase)
            Next

        End With

    End Sub
   
    Public Class clsPhaseDB
        Public AllRoles As List(Of clsRolleDB)
        Public AllCosts As List(Of clsKostenartDB)
        Public AllResults As List(Of clsResultDB)


        Public earliestStart As Integer
        Public latestStart As Integer
        Public minDauer As Integer
        Public maxDauer As Integer
        Public relStart As Integer
        Public relEnde As Integer
        Public startOffsetinDays As Integer
        Public dauerInDays As Integer
        Public name As String
        Public farbe As Object

        Public ReadOnly Property getResult(ByVal index As Integer) As clsResultDB

            Get
                getResult = AllResults.Item(index - 1)
            End Get

        End Property


        Sub copyFrom(ByVal phase As clsPhase, ByVal hfarbe As Object)
            Dim r As Integer, k As Integer

            With phase
                Me.earliestStart = .earliestStart
                Me.latestStart = .latestStart
                Me.minDauer = .minDauer
                Me.maxDauer = .maxDauer
                Me.relStart = .relStart
                Me.relEnde = .relEnde
                Me.startOffsetinDays = .startOffsetinDays
                Me.dauerInDays = .dauerInDays
                Me.name = .name
                Dim dimension As Integer

                ' Änderung 18.6 , weil Querschnittsphasen Namen jetzt der Projekt-Name ist ...
                Try
                    Me.farbe = .Farbe
                Catch ex As Exception
                    Me.farbe = hfarbe
                End Try


                For r = 1 To .CountRoles
                    'Dim newRole As New clsRolleDB(.relEnde - .relStart)
                    dimension = .getRole(r).getDimension
                    Dim newRole As New clsRolleDB(dimension)
                    newRole.copyFrom(.getRole(r))
                    AllRoles.Add(newRole)
                Next

                For r = 1 To .CountResults
                    Dim newResult As New clsResultDB

                    Try
                        newResult.CopyFrom(.getResult(r))
                        AllResults.Add(newResult)
                    Catch ex As Exception

                    End Try

                Next

                For k = 1 To .CountCosts
                    'Dim newCost As New clsKostenartDB(.relEnde - relStart)
                    dimension = .getCost(k).getDimension
                    Dim newCost As New clsKostenartDB(dimension)
                    newCost.copyFrom(.getCost(k))
                    AllCosts.Add(newCost)
                Next

            End With

        End Sub

        'Sub copyto(ByRef phase As clsPhase, ByVal ProjektStartdate As Date)

        '    Dim r As Integer, k As Integer

        '    With phase
        '        .earliestStart = Me.earliestStart
        '        .latestStart = Me.latestStart
        '        .minDauer = Me.minDauer
        '        .maxDauer = Me.maxDauer
        '        .relStart = Me.relStart
        '        .relEnde = Me.relEnde
        '        ' das Projektstartdatum muss mit übergeben werden, weil in dieser Methode
        '        ' die Werte für relstart und relende gesetzt werden 
        '        .startOffsetinDays = .startOffsetinDays
        '        .dauerInDays = .dauerInDays

        '        .name = Me.name

        '        For r = 1 To Me.AllRoles.Count
        '            Dim newRole As New clsRolle(.relEnde - .relStart)
        '            Me.AllRoles.Item(r - 1).copyto(newRole)
        '            .AddRole(newRole)

        '        Next

        '        Try
        '            Dim tstAnzahl As Integer = Me.AllResults.Count
        '            For r = 1 To tstAnzahl

        '                Dim newresult As New clsResult(parent:=phase)

        '                Try
        '                    Me.getResult(r).CopyTo(newresult)
        '                    .AddResult(newresult)
        '                Catch ex As Exception

        '                End Try

        '            Next
        '        Catch ex As Exception

        '        End Try



        '        For k = 1 To Me.AllCosts.Count
        '            Dim newCost As New clsKostenart(.relEnde - relStart)
        '            Me.AllCosts.Item(k - 1).copyto(newCost)
        '            .AddCost(newCost)
        '        Next

        '    End With


        'End Sub


        Sub copyto(ByRef phase As clsPhase)
            Dim r As Integer, k As Integer
            Dim dauer As Integer, startoffset As Integer

            With phase
                .earliestStart = Me.earliestStart
                .latestStart = Me.latestStart
                .minDauer = Me.minDauer
                .maxDauer = Me.maxDauer
                ' Änderung 28.11. relstart , relende ist nur noch readonly ; jetzt wird exaktes Datum mitgeführt
                '.relStart = Me.relStart
                '.relEnde = Me.relEnde
                .name = Me.name

                ' notwendig, da in älteren Versionen in der Datenbank evtl nur der wert für relende, relstart in Monaten gepeichert ist
                ' nicht aber der Wert für dauerindays oder startoffset
                If Me.dauerInDays = 0 Then
                    ' nutze 
                    startoffset = DateDiff(DateInterval.Day, .Parent.startDate, .Parent.startDate.AddMonths(Me.relStart - 1))
                    'dauer = DateDiff(DateInterval.Day, .Parent.startDate.AddMonths(Me.relStart - 1), .Parent.startDate.AddMonths(Me.relEnde).AddDays(-1)) + 1
                    dauer = calcDauerIndays(.Parent.startDate.AddDays(startoffset), Me.relEnde - Me.relStart + 1, True)
                Else
                    startoffset = Me.startOffsetinDays
                    dauer = Me.dauerInDays
                End If

                Dim dimension As Integer

                ' in der Datenbank wird es konsistent gespeichert
                For r = 1 To Me.AllRoles.Count
                    'Dim newRole As New clsRolle(.relEnde - .relStart)

                    dimension = Me.AllRoles.Item(r - 1).Bedarf.Length - 1
                    Dim newRole As New clsRolle(dimension)
                    Me.AllRoles.Item(r - 1).copyto(newRole)
                    .AddRole(newRole)

                Next

                For k = 1 To Me.AllCosts.Count
                    'Dim newCost As New clsKostenart(.relEnde - relStart)
                    dimension = Me.AllCosts.Item(k - 1).Bedarf.Length - 1
                    Dim newCost As New clsKostenart(dimension)
                    Me.AllCosts.Item(k - 1).copyto(newCost)
                    .AddCost(newCost)
                Next

                .changeStartandDauer(startoffset, dauer)

                Try
                    Dim tstAnzahl As Integer = Me.AllResults.Count
                    For r = 1 To tstAnzahl

                        Dim newresult As New clsResult(parent:=phase)

                        Try
                            Me.getResult(r).CopyTo(newresult)
                            .AddResult(newresult)
                        Catch ex As Exception

                        End Try

                    Next
                Catch ex As Exception

                End Try



                

            End With

        End Sub

        Sub New()
            AllRoles = New List(Of clsRolleDB)
            AllCosts = New List(Of clsKostenartDB)
            AllResults = New List(Of clsResultDB)
        End Sub
    End Class

    Public Class clsRolleDB

        Public RollenTyp As Integer
        Public name As String
        Public farbe As Object
        Public startkapa As Integer
        Public tagessatzIntern As Double
        Public tagessatzExtern As Double
        Public Bedarf() As Double
        Public isCalculated As Boolean

        Sub copyFrom(ByVal role As clsRolle)

            With role
                Me.RollenTyp = .RollenTyp
                Me.name = .name
                Me.farbe = .farbe
                Me.startkapa = .Startkapa
                Me.tagessatzIntern = .tagessatzIntern
                Me.tagessatzExtern = .tagessatzExtern
                Bedarf = .Xwerte
                Me.isCalculated = .isCalculated
            End With

        End Sub

        Sub copyto(ByRef role As clsRolle)

            With role
                .RollenTyp = Me.RollenTyp
                '.name = Me.name
                '.farbe = Me.farbe
                '.Startkapa = Me.startkapa
                '.tagessatzIntern = Me.tagessatzIntern
                '.tagessatzExtern
                .Xwerte = Me.Bedarf
                .isCalculated = Me.isCalculated
            End With

        End Sub

        Sub New()
            isCalculated = False
        End Sub

        Sub New(ByVal laenge As Integer)

            ReDim Bedarf(laenge)
            isCalculated = False

        End Sub

    End Class

    Public Class clsKostenartDB

        Public KostenTyp As Integer
        Public name As String
        Public farbe As Object
        Public Bedarf() As Double

        Sub copyFrom(ByVal cost As clsKostenart)

            With cost
                Me.KostenTyp = .KostenTyp
                Me.name = .name
                Me.farbe = .farbe
                Bedarf = .Xwerte
            End With

        End Sub

        Sub copyto(ByRef cost As clsKostenart)

            With cost
                .KostenTyp = Me.KostenTyp
                'Me.name = .name
                'Me.farbe = .farbe
                .Xwerte = Bedarf
            End With

        End Sub

        Sub New()

        End Sub

        Sub New(ByVal laenge As Integer)

            ReDim Bedarf(laenge)

        End Sub


    End Class

    Public Class clsResultDB

        Public bewertungen As SortedList(Of String, clsBewertungDB)


        Public name As String
        Public verantwortlich As String
        Public offset As Long

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


        Friend Sub CopyTo(ByRef newResult As clsResult)
            Dim i As Integer

            Try
                With newResult

                    .name = Me.name
                    .verantwortlich = Me.verantwortlich
                    .offset = Me.offset

                    For i = 1 To Me.bewertungsCount

                        Dim newb As New clsBewertung
                        Try
                            Me.getBewertung(i).copyto(newb)
                            .addBewertung(newb)
                        Catch ex1 As Exception

                        End Try

                    Next

                End With

            Catch ex As Exception

            End Try

            

        End Sub

        Friend Sub CopyFrom(ByVal newResult As clsResult)
            Dim i As Integer
            Dim newb As New clsBewertungDB

            With newResult

                Me.name = .name
                Me.verantwortlich = .verantwortlich
                Me.offset = .offset

                Try
                    For i = 1 To .bewertungsCount
                        newb.copyfrom(.getBewertung(i))
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

            bewertungen = New SortedList(Of String, clsBewertungDB)

        End Sub


    End Class

    Public Class clsBewertungDB

        Public color As Integer
        Public description As String
        Public bewerterName As String
        Public datum As Date

        Friend Sub CopyTo(ByRef newB As clsBewertung)

            With newB
                .colorIndex = Me.color
                .description = Me.description
                .datum = Me.datum
                .bewerterName = Me.bewerterName
            End With

        End Sub

        Friend Sub Copyfrom(ByVal b As clsBewertung)

            Me.color = b.colorIndex
            Me.description = b.description
            Me.bewerterName = b.bewerterName
            Me.datum = b.datum

        End Sub

        Sub New()
            bewerterName = ""
            datum = Nothing
            color = 0
            description = ""
        End Sub

    End Class


    Public Sub New()

        AllPhases = New List(Of clsPhaseDB)

    End Sub

End Class
