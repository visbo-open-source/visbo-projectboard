''' <summary>
''' Vorsicht !!! 
''' bei allen Änderungen in clsProjektDB und in clsPhaseDB, da für den ReST-Server-Zugriff separate Klassen existieren, die aber fast gleich sind.
''' 
''' Klassen-Definition für eine Phase bzw Sammel-Task in MongoDB
''' </summary>
''' <remarks></remarks>
Public Class clsPhaseDB
    Public AllRoles As List(Of clsRolleDB)
    Public AllCosts As List(Of clsKostenartDB)
    Public AllResults As List(Of clsResultDB)
    Public AllBewertungen As SortedList(Of String, clsBewertungDB)

    Public percentDone As Double
    Public responsible As String

    ' Ergänzungen 8.5.18 wegen Dokumenten Urls
    Public docURL As String
    Public docUrlAppID As String

    Public deliverables As List(Of String)

    Public ampelStatus As Integer
    Public ampelErlaeuterung As String

    Public earliestStart As Integer
    Public latestStart As Integer
    Public minDauer As Integer
    Public maxDauer As Integer
    Public relStart As Integer
    Public relEnde As Integer
    Public startOffsetinDays As Integer
    Public dauerInDays As Integer
    Public name As String
    Public farbe As Integer

    Public shortName As String
    Public originalName As String
    Public appearance As String

    Public ReadOnly Property getMilestone(ByVal index As Integer) As clsResultDB

        Get
            getMilestone = AllResults.Item(index - 1)
        End Get

    End Property


    Sub copyFrom(ByVal phase As clsPhase, ByVal hfarbe As Integer)

        Dim i As Integer, r As Integer, k As Integer

        With phase
            Me.earliestStart = .earliestStart
            Me.latestStart = .latestStart
            'Me.minDauer = .minDauer
            'Me.maxDauer = .maxDauer
            Me.relStart = .relStart
            Me.relEnde = .relEnde
            Me.startOffsetinDays = .startOffsetinDays
            Me.dauerInDays = .dauerInDays
            Me.name = .nameID

            Me.shortName = .shortName
            Me.originalName = .originalName
            Me.appearance = .appearance

            ' Dokumenten-Url und Applikation
            Me.docURL = .DocURL
            Me.docUrlAppID = .DocUrlAppID

            Me.responsible = .verantwortlich
            Me.percentDone = .percentDone

            Me.ampelErlaeuterung = .ampelErlaeuterung
            Me.ampelStatus = .ampelStatus

            Dim dimension As Integer

            ' Änderung 18.6 , ab 29.5 .16 kann jeder Phase auch eine Farbe zugewiesen werden 
            Try
                Me.farbe = .individualColor
            Catch ex As Exception
                Me.farbe = hfarbe
            End Try

            ' jetzt die evtl vorhandenen Deliverables zuweisen ...
            For i = 1 To .countDeliverables
                Dim tmpDeliverable As String = .getDeliverable(i)
                Me.deliverables.Add(tmpDeliverable)
            Next


            For r = 1 To .countRoles
                'Dim newRole As New clsRolleDB(.relEnde - .relStart)
                dimension = .getRole(r).getDimension
                Dim newRole As New clsRolleDB(dimension)
                newRole.copyFrom(.getRole(r))
                AllRoles.Add(newRole)
            Next

            For r = 1 To .countMilestones
                Dim newResult As New clsResultDB

                Try
                    newResult.CopyFrom(.getMilestone(r))
                    AllResults.Add(newResult)
                Catch ex As Exception

                End Try

            Next

            For k = 1 To .countCosts
                'Dim newCost As New clsKostenartDB(.relEnde - relStart)
                dimension = .getCost(k).getDimension
                Dim newCost As New clsKostenartDB(dimension)
                newCost.copyFrom(.getCost(k))
                AllCosts.Add(newCost)
            Next


            ' jetzt evtl vorhandene Bewertungen abspeichern ... 
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


    Sub copyto(ByRef phase As clsPhase, Optional phaseNr As Integer = 100)
        Dim r As Integer, k As Integer
        Dim dauer As Integer, startoffset As Integer

        With phase
            .earliestStart = Me.earliestStart
            .latestStart = Me.latestStart
            '.minDauer = Me.minDauer
            '.maxDauer = Me.maxDauer

            ' jetzt kommen die Doumenten Folder und AppIDs
            If Not IsNothing(Me.docURL) Then
                .DocURL = Me.docURL
            End If

            If Not IsNothing(Me.docUrlAppID) Then
                .DocUrlAppID = Me.docUrlAppID
            End If

            If Not IsNothing(Me.responsible) Then
                .verantwortlich = Me.responsible
            End If

            If Not IsNothing(Me.percentDone) Then
                .percentDone = Me.percentDone
            End If

            If Not IsNothing(Me.shortName) Then
                .shortName = Me.shortName
            End If

            If Not IsNothing(Me.originalName) Then
                .originalName = Me.originalName
            End If

            If Not IsNothing(Me.appearance) Then
                .appearance = Me.appearance
            End If

            ' jetzt die Deliverables aufnehmen 
            If Not IsNothing(Me.deliverables) Then
                If Me.deliverables.Count > 0 Then
                    For i = 1 To Me.deliverables.Count
                        Dim tmpDeliverable As String = Me.deliverables.Item(i - 1)
                        .addDeliverable(tmpDeliverable)
                    Next
                End If
            End If

            ' Ergänzung 9.5.16 AmpelStatus und Erläuterung mitaufgenommen ... 
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung

            Try
                .farbe = Me.farbe
            Catch ex As Exception

            End Try


            ' Änderung tk 20.4.2015
            ' damit alte Datenbank Einträge ohne Hierarchie auch noch gelesen werden können ..
            If Not istElemID(Me.name) Then
                If phaseNr = 1 Then
                    .nameID = rootPhaseName
                Else
                    .nameID = calcHryElemKey(Me.name, False)
                End If

            Else
                .nameID = Me.name
            End If


            ' notwendig, da in älteren Versionen in der Datenbank evtl nur der wert für relende, relstart in Monaten gepeichert ist
            ' nicht aber der Wert für dauerindays oder startoffset
            If Me.dauerInDays = 0 Then
                ' nutze 
                startoffset = CInt(DateDiff(DateInterval.Day, .parentProject.startDate, .parentProject.startDate.AddMonths(Me.relStart - 1)))
                'dauer = DateDiff(DateInterval.Day, .Parent.startDate.AddMonths(Me.relStart - 1), .Parent.startDate.AddMonths(Me.relEnde).AddDays(-1)) + 1
                dauer = calcDauerIndays(.parentProject.startDate.AddDays(startoffset), Me.relEnde - Me.relStart + 1, True)
            Else
                startoffset = Me.startOffsetinDays
                dauer = Me.dauerInDays
            End If

            Dim dimension As Integer

            ' macht nur Sinn, die Rollen zu holen, wenn auch Rollen-Definitionen vorhanden isnd
            ' im PPTSmartInfo sindsie es noch nicht ! 
            If RoleDefinitions.Count > 0 Then
                ' nur aufrufen, wenn 
                For r = 1 To Me.AllRoles.Count
                    'Dim newRole As New clsRolle(.relEnde - .relStart)

                    dimension = Me.AllRoles.Item(r - 1).Bedarf.Length - 1
                    Dim newRole As New clsRolle(dimension)
                    Me.AllRoles.Item(r - 1).copyto(newRole)
                    .addRole(newRole)

                Next
            End If

            ' macht nur Sinn, die Rollen zu holen, wenn auch Rollen-Definitionen vorhanden isnd
            ' im PPTSmartInfo sind sie es noch nicht ! 
            If CostDefinitions.Count > 0 Then
                For k = 1 To Me.AllCosts.Count
                    'Dim newCost As New clsKostenart(.relEnde - relStart)
                    dimension = Me.AllCosts.Item(k - 1).Bedarf.Length - 1
                    Dim newCost As New clsKostenart(dimension)
                    Me.AllCosts.Item(k - 1).copyto(newCost)
                    .AddCost(newCost)
                Next
            End If

            .changeStartandDauer(startoffset, dauer)

            Try
                Dim tstAnzahl As Integer = Me.AllResults.Count
                For r = 1 To tstAnzahl

                    Dim newresult As New clsMeilenstein(parent:=phase)

                    Try
                        Me.getMilestone(r).CopyTo(newresult)
                        .addMilestone(newresult)
                    Catch ex As Exception

                    End Try

                Next
            Catch ex As Exception

            End Try

            ' evtl vorhandene Bewertungen kopieren .... 
            For b As Integer = 1 To Me.AllBewertungen.Count
                Dim newb As New clsBewertung
                Me.AllBewertungen.ElementAt(b - 1).Value.CopyTo(newb)

                Try
                    .addBewertung(newb)
                Catch ex As Exception

                End Try

            Next




        End With

    End Sub

    Friend Sub addBewertung(ByVal b As clsBewertungDB)
        Dim key As String

        If Not IsNothing(b.bewerterName) Then
            key = b.bewerterName.Trim & "#" & b.datum.ToString("MMM yy")
        Else
            key = "#" & b.datum.ToString("MMM yy")
        End If

        Try
            Me.AllBewertungen.Add(key, b)
        Catch ex As Exception

            Throw New ArgumentException("Bewertung wurde bereits vergeben ..")

        End Try

    End Sub

    Sub New()
        AllRoles = New List(Of clsRolleDB)
        AllCosts = New List(Of clsKostenartDB)
        AllResults = New List(Of clsResultDB)
        AllBewertungen = New SortedList(Of String, clsBewertungDB)

        deliverables = New List(Of String)

        percentDone = 0.0
        responsible = ""

        ampelStatus = 0
        ampelErlaeuterung = ""

        shortName = ""
        originalName = ""
        appearance = ""

        docURL = ""
        docUrlAppID = ""

    End Sub
End Class
