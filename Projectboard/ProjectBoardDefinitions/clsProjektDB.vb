Public Class clsProjektDB

    Public name As String
    Public variantName As String
    Public Risiko As Double
    Public StrategicFit As Double

    ' Änderung tk: die CustomFields ergänzt ...
    Public customDblFields As SortedList(Of String, Double)
    Public customStringFields As SortedList(Of String, String)
    Public customBoolFields As SortedList(Of String, Boolean)

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
    Public farbe As Integer
    Public Schrift As Integer
    Public Schriftfarbe As Object
    Public VorlagenName As String
    Public Dauer As Integer
    Public AllPhases As List(Of clsPhaseDB)
    Public hierarchy As clsHierarchyDB
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

            ' wenn es einen Varianten-Namen gibt, wird als Datenbank Name 
            ' .name = calcprojektkey(projekt) abgespeichert; das macht das Auslesen später effizienter 

            Me.name = calcProjektKeyDB(projekt.name, projekt.variantName)

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
            Me.Dauer = .anzahlRasterElemente
            ' ergänzt am 16.11.13
            Me.volumen = .volume
            Me.complexity = .complexity
            Me.description = .description
            Me.businessUnit = .businessUnit

            Me.hierarchy.copyFrom(projekt.hierarchy)

            For i = 1 To .CountPhases
                Dim newPhase As New clsPhaseDB
                newPhase.copyFrom(.getPhase(i), .farbe)
                AllPhases.Add(newPhase)
            Next

            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 
            For Each kvp As KeyValuePair(Of Integer, String) In projekt.customStringFields
                Me.customStringFields.Add(CStr(kvp.Key), kvp.Value)
            Next

            For Each kvp As KeyValuePair(Of Integer, Double) In projekt.customDblFields
                Me.customDblFields.Add(CStr(kvp.Key), kvp.Value)
            Next

            For Each kvp As KeyValuePair(Of Integer, Boolean) In projekt.customBoolFields
                Me.customBoolFields.Add(CStr(kvp.Key), kvp.Value)
            Next


        End With

    End Sub

    Public Sub copyto(ByRef projekt As clsProjekt)
        Dim i As Integer
        Dim tmpstr(5) As String


        With projekt
            .timeStamp = Me.timestamp.ToLocalTime
            .Id = Me.Id

            ' jetzt muss der Datenbank Name aufgesplittet werden in name und variant-Name
            If Me.variantName <> "" And Me.variantName.Trim.Length > 0 Then
                tmpstr = Me.name.Split(New Char() {CChar("#")}, 3)
                If tmpstr.Length > 1 Then
                    If tmpstr(1) = Me.variantName Then
                        .name = tmpstr(0)
                    Else
                        .name = Me.name
                    End If
                Else
                    .name = Me.name
                End If
            Else
                .name = Me.name
            End If

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

            ' Änderung 18.5.2014: jetzt prüfen, ob diese Vorlage existiert: 
            ' wenn ja, dann übernehmen Farbe, Schrift und Schriftfarbe
            Try
                If Projektvorlagen.Contains(.VorlagenName) Then
                    Dim pvorlage As clsProjektvorlage = Projektvorlagen.getProject(.VorlagenName)
                    .Schrift = pvorlage.Schrift
                    .Schriftfarbe = pvorlage.Schriftfarbe
                    .farbe = pvorlage.farbe
                End If
            Catch ex As Exception

            End Try

            Me.hierarchy.copyTo(projekt.hierarchy)

            '.Dauer = Me.Dauer
            For i = 1 To Me.AllPhases.Count
                Dim newPhase As New clsPhase(projekt)
                AllPhases.Item(i - 1).copyto(newPhase, i)
                .AddPhase(newPhase)
            Next

            ' jetzt werden Ampel Status und Beschreibung gesetzt 
            ' da das jetzt in der Phase(1) abgespeichert ist, darf das erst gemacht werden, wenn die Phasen alle kopiert sind ... 
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung

            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 

            If Not IsNothing(Me.customStringFields) Then
                For Each kvp As KeyValuePair(Of String, String) In Me.customStringFields
                    projekt.customStringFields.Add(CInt(kvp.Key), kvp.Value)
                Next
            End If
            If Not IsNothing(Me.customDblFields) Then
                For Each kvp As KeyValuePair(Of String, Double) In Me.customDblFields
                    projekt.customDblFields.Add(CInt(kvp.Key), kvp.Value)
                Next
            End If
            If Not IsNothing(Me.customBoolFields) Then
                For Each kvp As KeyValuePair(Of String, Boolean) In Me.customBoolFields
                    projekt.customBoolFields.Add(CInt(kvp.Key), kvp.Value)
                Next
            End If
            

        End With

    End Sub

    Public Class clsHierarchyDB
        Public allNodes As SortedList(Of String, clsHierarchyNodeDB)

        ''' <summary>
        ''' kopiert aus einem HSP-Element in ein DB-Element
        ''' </summary>
        ''' <param name="hry"></param>
        ''' <remarks></remarks>
        Sub copyFrom(ByVal hry As clsHierarchy)

            Dim hryNode As clsHierarchyNode
            Dim elemID As String
            Dim hryNodeDB As clsHierarchyNodeDB

            For i = 1 To hry.count

                hryNodeDB = New clsHierarchyNodeDB

                elemID = hry.getIDAtIndex(i)
                If elemID = rootPhaseName Then
                    elemID = rootPhaseNameDB
                End If
                If elemID.Contains(punktName) Then
                    elemID = elemID.Replace(punktName, punktNameDB)
                End If
                hryNode = hry.nodeItem(i)
                hryNodeDB.copyFrom(hryNode)

                Me.allNodes.Add(elemID, hryNodeDB)

            Next

        End Sub

        ''' <summary>
        ''' kopiert aus einem DB Element in ein HSP Element 
        ''' </summary>
        ''' <param name="hry"></param>
        ''' <remarks></remarks>
        Sub copyTo(ByRef hry As clsHierarchy)

            Dim hryNode As clsHierarchyNode
            Dim elemID As String
            Dim hryNodeDB As clsHierarchyNodeDB

            For i = 1 To Me.allNodes.Count

                hryNode = New clsHierarchyNode

                elemID = Me.allNodes.ElementAt(i - 1).Key
                If elemID = rootPhaseNameDB Then
                    elemID = rootPhaseName
                End If
                If elemID.Contains(punktNameDB) Then
                    elemID = elemID.Replace(punktNameDB, punktName)
                End If
                hryNodeDB = Me.allNodes.ElementAt(i - 1).Value
                hryNodeDB.copyTo(hryNode)

                hry.copyNode(hryNode, elemID)

            Next

        End Sub

        Sub New()
            allNodes = New SortedList(Of String, clsHierarchyNodeDB)
        End Sub
    End Class

    Public Class clsHierarchyNodeDB
        Public elemName As String
        Public origName As String
        Public indexOfElem As Integer
        Public parentNodeKey As String
        Public childNodeKeys As List(Of String)

        ' 
        ''' <summary>
        ''' kopiert einen HAuptspeicher Hierarchie Knoten in einen DB Hierarchie Knoten 
        ''' </summary>
        ''' <param name="hryNode"></param>
        ''' <remarks></remarks>
        Sub copyFrom(ByVal hryNode As clsHierarchyNode)

            Dim childID As String
            With hryNode
                Me.elemName = .elemName
                ' ist seit 29.5 niht mehr Bestandteil eines Hierarchie Knotens
                'Me.origName = .origName
                Me.indexOfElem = .indexOfElem
                Me.parentNodeKey = .parentNodeKey
                For i As Integer = 1 To .childCount
                    childID = .getChild(i)
                    Me.childNodeKeys.Add(childID)
                Next
            End With

        End Sub

        ''' <summary>
        ''' kopiert einen DB Hierarchie-Knoten in einen Hauptspeicher Hierarchie Knoten 
        ''' </summary>
        ''' <param name="hryNode"></param>
        ''' <remarks></remarks>
        Sub copyTo(ByRef hryNode As clsHierarchyNode)

            Dim childID As String
            With hryNode
                .elemName = Me.elemName
                ' ist seit 29.5 nicht mehr Bestandteil eines Hierarchie-Knotens 
                '.origName = Me.origName
                .indexOfElem = Me.indexOfElem
                .parentNodeKey = Me.parentNodeKey
                For i As Integer = 1 To Me.childNodeKeys.Count
                    childID = Me.childNodeKeys.Item(i - 1)
                    .addChild(childID)
                Next
            End With

        End Sub

        Sub New()

            childNodeKeys = New List(Of String)

        End Sub

    End Class

    Public Class clsPhaseDB
        Public AllRoles As List(Of clsRolleDB)
        Public AllCosts As List(Of clsKostenartDB)
        Public AllResults As List(Of clsResultDB)

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

            Dim r As Integer, k As Integer

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

                Me.ampelErlaeuterung = .ampelErlaeuterung
                Me.ampelStatus = .ampelStatus

                Dim dimension As Integer

                ' Änderung 18.6 , ab 29.5 .16 kann jeder Phase auch eine Farbe zugewiesen werden 
                Try
                    Me.farbe = .individualColor
                Catch ex As Exception
                    Me.farbe = hfarbe
                End Try


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

                If Not IsNothing(Me.shortName) Then
                    .shortName = Me.shortName
                End If

                If Not IsNothing(Me.originalName) Then
                    .originalName = Me.originalName
                End If

                If Not IsNothing(Me.appearance) Then
                    .appearance = Me.appearance
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

                ' in der Datenbank wird es konsistent gespeichert
                For r = 1 To Me.AllRoles.Count
                    'Dim newRole As New clsRolle(.relEnde - .relStart)

                    dimension = Me.AllRoles.Item(r - 1).Bedarf.Length - 1
                    Dim newRole As New clsRolle(dimension)
                    Me.AllRoles.Item(r - 1).copyto(newRole)
                    .addRole(newRole)

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

                        Dim newresult As New clsMeilenstein(parent:=phase)

                        Try
                            Me.getMilestone(r).CopyTo(newresult)
                            .addMilestone(newresult)
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

            ampelStatus = 0
            ampelErlaeuterung = ""

            shortName = ""
            originalName = ""
            appearance = ""

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
                Me.startkapa = CInt(.Startkapa)
                Me.tagessatzIntern = .tagessatzIntern
                Me.tagessatzExtern = .tagessatzExtern
                Bedarf = .Xwerte
                Me.isCalculated = .isCalculated
            End With

        End Sub

        Sub copyto(ByRef role As clsRolle)

            With role
                .RollenTyp = Me.RollenTyp
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
        Public alternativeColor As Long

        Public shortName As String
        Public originalName As String
        Public appearance As String

        Public deliverables As List(Of String)

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

                    Try
                        If Not IsNothing(Me.alternativeColor) Then
                            .farbe = CInt(Me.alternativeColor)
                        Else
                            .farbe = CInt(awinSettings.AmpelNichtBewertet)
                        End If
                    Catch ex As Exception

                    End Try

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

                For i = 1 To .countDeliverables
                    Dim tmpDeliverable As String = .getDeliverable(i)
                    Me.deliverables.Add(tmpDeliverable)
                Next

                Try
                    For i = 1 To .bewertungsCount
                        Dim newb As New clsBewertungDB
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
            deliverables = New List(Of String)

        End Sub


    End Class


    Public Class clsBewertungDB
        ' Änderung tk: 2.11 deliverables / Ergebnisse hinzugefügt 

        Public color As Integer
        Public description As String
        Public deliverables As String
        Public bewerterName As String
        Public datum As Date

        Friend Sub CopyTo(ByRef newB As clsBewertung)

            With newB
                .colorIndex = Me.color
                .description = Me.description
                '.deliverables = Me.deliverables
                .datum = Me.datum
                .bewerterName = Me.bewerterName
            End With

        End Sub

        Friend Sub Copyfrom(ByVal b As clsBewertung)

            Me.color = b.colorIndex
            Me.description = b.description
            'Me.deliverables = b.deliverables
            Me.bewerterName = b.bewerterName
            Me.datum = b.datum

        End Sub

        Sub New()
            bewerterName = ""
            datum = Nothing
            color = 0
            description = ""
            deliverables = ""
        End Sub

    End Class


    Public Sub New()

        AllPhases = New List(Of clsPhaseDB)
        hierarchy = New clsHierarchyDB

        customDblFields = New SortedList(Of String, Double)
        customStringFields = New SortedList(Of String, String)
        customBoolFields = New SortedList(Of String, Boolean)

    End Sub

End Class
