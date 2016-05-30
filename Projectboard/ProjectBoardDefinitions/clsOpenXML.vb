''' <summary>
''' diese Klasse dient zum Austausch von VISBO PRojekten 
''' </summary>
''' <remarks></remarks>
Public Class clsOpenXML

    Public projectName As String
    Public variantName As String
    Public timeStamp As Date
    Public ID As String

    Public projectType As String

    Public sourceDBURL As String
    Public sourceDBName As String
    Public sourceUID As String

    Public projectStakeholder As List(Of String)

    Public budget As Double
    Public currency As String

    Public projectTitle As String
    Public status As String
    Public businessUnit As String

    Public strategicFit As Double
    Public risk As Double

    ' die Custom Fields für ein Projekt 
    Public customDblFields As SortedList(Of Integer, Double)
    Public customStringFields As SortedList(Of Integer, String)
    Public customBoolFields As SortedList(Of Integer, Boolean)

    Public tasks As List(Of clsOpenTask)


    Public Sub copyFrom(ByVal projekt As clsProjekt)


        With projekt
            Me.projectName = .name
            Me.variantName = .variantName

            If Not IsNothing(.timeStamp) Then
                Me.timeStamp = .timeStamp.ToUniversalTime
            Else
                Me.timeStamp = Date.UtcNow
            End If

            ' die folgenden Infos werden noch nicht besetzt 
            Me.sourceDBURL = ""
            Me.sourceDBName = ""

            If Not IsNothing(.Id) Then
                Me.sourceUID = .Id
            End If

            Me.projectType = .VorlagenName

            Me.budget = .Erloes
            Me.currency = "€"

            Me.projectTitle = ""
            Me.status = .Status
            Me.businessUnit = .businessUnit

            Me.strategicFit = .StrategicFit
            Me.risk = .Risiko

            ' jetzt werden die CustomFields bestückt .... 
            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 
            For Each kvp As KeyValuePair(Of Integer, String) In projekt.customStringFields
                Me.customStringFields.Add(kvp.Key, kvp.Value)
            Next

            For Each kvp As KeyValuePair(Of Integer, Double) In projekt.customDblFields
                Me.customDblFields.Add(kvp.Key, kvp.Value)
            Next

            For Each kvp As KeyValuePair(Of Integer, Boolean) In projekt.customBoolFields
                Me.customBoolFields.Add(kvp.Key, kvp.Value)
            Next



            ' jetzt müssen alle Phasen kopiert werden , dabei enthält die erste Phase ja viel 
            ' Projekt-Information 

            ' hier soll er sich leiten lassen von der Hierarchie 
            ' das Projekt selber, das heisst phase(1) hat wbsCode ""
            Dim pspCode As String = ""


            For i As Integer = 1 To .CountPhases

                Dim newPhase As New clsOpenTask

                If i = 1 Then
                    ' es ist die erste Phase, also die Projekt-Phase ...
                    newPhase.copyFrom(.getPhase(i), .leadPerson, .description)
                Else
                    ' es ist nicht die erste Phase 
                    newPhase.copyFrom(.getPhase(i))
                End If

                tasks.Add(newPhase)

            Next


        End With


    End Sub

    Public Sub copyTo(ByRef projekt As clsProjekt)

        Dim i As Integer
        Dim tmpstr(5) As String


        With projekt

            Dim task0 As clsOpenTask = Me.tasks.Item(0)

            .name = Me.projectName
            .variantName = Me.variantName
            .timeStamp = Me.timeStamp.ToLocalTime
            .Id = Me.ID
            .Risiko = Me.risk
            .StrategicFit = Me.strategicFit
            .Erloes = Me.budget
            .leadPerson = task0.responsible
            .tfZeile = 0
            .startDate = task0.startDate.ToLocalTime

            .earliestStart = task0.earliestStartOffset
            .latestStart = task0.latestStartOffset

            If .earliestStart < 0 Then
                .earliestStartDate = .startDate.AddMonths(.earliestStart)
            Else
                .earliestStart = 0
            End If

            If .latestStart > 0 Then
                .latestStartDate = .startDate.AddMonths(.latestStart)
            Else
                .latestStart = 0
            End If

            ' .projekttitle = Me.projekttitle
            .Status = Me.status
            .businessUnit = Me.businessUnit
            .description = task0.description

            .VorlagenName = Me.projectType
            Try
                Dim pvorlage As clsProjektvorlage
                If Projektvorlagen.Contains(.VorlagenName) Then
                    pvorlage = Projektvorlagen.getProject(.VorlagenName)

                Else
                    pvorlage = Projektvorlagen.getProject(0)
                    .Schrift = pvorlage.Schrift
                    .Schriftfarbe = pvorlage.Schriftfarbe
                    .farbe = pvorlage.farbe
                End If
                .Schrift = pvorlage.Schrift
                .Schriftfarbe = pvorlage.Schriftfarbe
                .farbe = pvorlage.farbe

            Catch ex As Exception
                .Schrift = 10
                .Schriftfarbe = RGB(0, 0, 0)
                .farbe = RGB(110, 110, 100)
            End Try


            ' jetzt werden die CustomFields rausgeschrieben, so fern es welche gibt ... 

            If Not IsNothing(Me.customStringFields) Then
                For Each kvp As KeyValuePair(Of Integer, String) In Me.customStringFields
                    projekt.customStringFields.Add(kvp.Key, kvp.Value)
                Next
            End If


            If Not IsNothing(Me.customDblFields) Then
                For Each kvp As KeyValuePair(Of Integer, Double) In Me.customDblFields
                    projekt.customDblFields.Add(kvp.Key, kvp.Value)
                Next
            End If


            If Not IsNothing(Me.customBoolFields) Then
                For Each kvp As KeyValuePair(Of Integer, Boolean) In Me.customBoolFields
                    projekt.customBoolFields.Add(kvp.Key, kvp.Value)
                Next
            End If

            For i = 1 To Me.tasks.Count
                Dim newPhase As New clsPhase(parent:=projekt)
                tasks.Item(i - 1).copyTo(newPhase, task0.startDate.ToLocalTime, i)
                .AddPhase(newPhase)
            Next


            ' jetzt muss die Ampel und Erläuterung noch geschrieben werden 
            Try
                .ampelStatus = .getPhase(1).ampelStatus
                .ampelErlaeuterung = .getPhase(1).ampelErlaeuterung
            Catch ex As Exception

            End Try
            

        End With

    End Sub

    ''' <summary>
    ''' Konstruktor für ein neues Projekt
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        projectName = "Testproject"
        variantName = ""
        timeStamp = Date.Now
        ID = ""

        projectType = ""
        sourceDBURL = ""
        sourceDBName = ""
        sourceUID = ""

        projectStakeholder = New List(Of String)

        budget = 0.0
        currency = "€"

        projectTitle = ""
        status = ""
        businessUnit = ""

        strategicFit = 5
        risk = 5

        customDblFields = New SortedList(Of Integer, Double)
        customStringFields = New SortedList(Of Integer, String)
        customBoolFields = New SortedList(Of Integer, Boolean)


        tasks = New List(Of clsOpenTask)

    End Sub
    '
    ' ############################ Klasse clsOpenTask
    '
    ''' <summary>
    ''' Klasse Phase
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsOpenTask

        Public name As String
        Public originalName As String
        Public abbreviation As String
        Public appearance As String
        Public color As Integer
        Public breadCrumb As String

        Public sourceUID As String

        Public description As String
        Public risks As List(Of clsOpenRiskChance)

        Public startDate As Date
        Public finishDate As Date
        Public earliestStartOffset As Integer
        Public latestStartOffset As Integer

        Public responsible As String

        Public ratings As List(Of clsOpenRating)

        Public costNeeds As List(Of clsOpenCostNeed)
        Public resourceNeeds As List(Of clsOpenResourceNeed)

        Public milestones As List(Of clsOpenMilestone)

        Public Sub copyFrom(ByVal cPhase As clsPhase, _
                            Optional ByVal optResponsible As String = Nothing, _
                            Optional ByVal optDescription As String = Nothing)

            Dim r As Integer, k As Integer
            Dim dimension As Integer

            With cPhase
                Me.name = .name
                Me.originalName = .originalName
                Me.abbreviation = .shortName
                Me.appearance = .appearance
                Me.color = .farbe
                Me.breadCrumb = .parentProject.hierarchy.getBreadCrumb(.nameID)

                'Me.sourceUID = ""

                If Not IsNothing(optDescription) Then
                    Me.description = optDescription
                Else
                    Me.description = ""
                End If

                ' keine Risks bisher definiert ...

                Me.startDate = .getStartDate.ToUniversalTime
                Me.finishDate = .getEndDate.ToUniversalTime
                Me.earliestStartOffset = .earliestStart
                Me.latestStartOffset = .latestStart

                If Not IsNothing(optResponsible) Then
                    Me.responsible = optResponsible
                Else
                    Me.responsible = ""
                End If

                For i = 1 To .bewertungsCount
                    Dim newRating As New clsOpenRating
                    newRating.description = .getBewertung(i).description
                    newRating.color = .getBewertung(i).colorIndex
                    newRating.rater = .getBewertung(i).bewerterName
                    newRating.ratingDate = .getBewertung(i).datum
                    Me.ratings.Add(newRating)
                Next


                For r = 1 To .countRoles
                    'Dim newRole As New clsRolleDB(.relEnde - .relStart)
                    dimension = .getRole(r).getDimension
                    Dim newOpenRole As New clsOpenResourceNeed(dimension)
                    newOpenRole.copyFrom(.getRole(r))
                    resourceNeeds.Add(newOpenRole)
                Next

                For k = 1 To .countCosts - 1
                    dimension = .getCost(k).getDimension
                    Dim newCost As New clsOpenCostNeed(dimension)
                    newCost.copyFrom(.getCost(k))
                    costNeeds.Add(newCost)
                Next

                For r = 1 To .countMilestones
                    Dim newOpenMilestone As New clsOpenMilestone

                    Try
                        newOpenMilestone.copyFrom(.getMilestone(r))
                        milestones.Add(newOpenMilestone)
                    Catch ex As Exception

                    End Try

                Next



            End With


        End Sub

        Public Sub copyTo(ByRef phase As clsPhase, ByVal projectStart As Date, _
                          Optional phaseNr As Integer = 100)


            Dim r As Integer, k As Integer
            'Dim dauer As Integer, startoffset As Integer

            With phase
                .earliestStart = Me.earliestStartOffset
                .latestStart = Me.latestStartOffset

                ' Ergänzung 9.5.16 AmpelStatus und Erläuterung mitaufgenommen ... 
                For i = 1 To Me.ratings.Count
                    Dim newb As New clsBewertung
                    Me.ratings.Item(i - 1).copyTo(newb)
                    .addBewertung(newb)
                Next


                Try
                    .farbe = Me.color
                Catch ex As Exception

                End Try

                Dim tmpDauer As Long = 0
                Dim phaseStartOffset As Long = 0


                If phaseNr = 1 Then
                    .nameID = rootPhaseName
                    phaseStartOffset = 0
                    tmpDauer = calcDauerIndays(Me.startDate.ToLocalTime, Me.finishDate.ToLocalTime)

                Else
                    .nameID = .parentProject.hierarchy.findUniqueElemKey(Me.name, False)
                    phaseStartOffset = DateDiff(DateInterval.Day, projectStart, Me.startDate.ToLocalTime)
                    tmpDauer = calcDauerIndays(Me.startDate.ToLocalTime, Me.finishDate.ToLocalTime)

                End If

                .changeStartandDauer(phaseStartOffset, tmpDauer)

                Dim anzahlMonate As Integer = getColumnOfDate(Me.finishDate.ToLocalTime) - getColumnOfDate(Me.startDate.ToLocalTime) + 1


                ' jetzt wird ausgelesen: stimmt die Dimension ? 
                ' wenn nein, wird einfach die Summe hergenommen und verteilt ... 
                Dim Xwerte() As Double
                'Dim oldWerte() As Double
                Dim roleUID As Integer


                ' die Rollen und Ressourcenbedarfe aufnehmen ...
                For r = 1 To Me.resourceNeeds.Count
                    Dim roleXML As clsOpenResourceNeed = Me.resourceNeeds.Item(r - 1)

                    Dim dimension As Integer = roleXML.monthlyNeeds.Length - 1
                    Dim newRole As New clsRolle(anzahlMonate - 1)

                    If dimension = anzahlMonate - 1 And roleXML.monthlyNeeds.Sum > 0 Then
                        ' alles in Ordnung , die Länge passt ...
                    Else
                        ' einfach die Summe hernehmen und verteilen ...
                        'ReDim oldWerte(0)
                        'oldWerte(0) = roleXML.sum
                        ReDim Xwerte(anzahlMonate - 1)
                        Call .berechneBedarfe(Me.startDate.ToLocalTime, Me.finishDate.ToLocalTime, roleXML.monthlyNeeds, 1.0, Xwerte)
                        roleXML.monthlyNeeds = Xwerte
                    End If

                    If RoleDefinitions.containsName(roleXML.resourceName) Then
                        roleUID = CInt(RoleDefinitions.getRoledef(roleXML.resourceName).UID)
                    Else
                        ' Rolle existiert noch nicht
                        ' wird hier neu aufgenommen

                        Dim newRoleDef As New clsRollenDefinition
                        newRoleDef.name = roleXML.resourceName
                        newRoleDef.farbe = RGB(120, 120, 120)
                        newRoleDef.Startkapa = 20

                        ' OvertimeRate in Tagessatz umrechnen
                        newRoleDef.tagessatzExtern = 780

                        ' StandardRate in Tagessatz umrechnen
                        newRoleDef.tagessatzIntern = 780

                        newRoleDef.UID = RoleDefinitions.Count + 1
                        If Not missingRoleDefinitions.containsName(newRoleDef.name) Then
                            missingRoleDefinitions.Add(newRoleDef)
                        End If

                        RoleDefinitions.Add(newRoleDef)

                        roleUID = newRoleDef.UID
                    End If


                    newRole.RollenTyp = roleUID
                    newRole.Xwerte = roleXML.monthlyNeeds

                    ' jetzt zur Phase dazu tun 
                    .addRole(newRole)

                Next

                Dim costUID As Integer

                ' die Kostenbedarfe aufnehmen ...
                For k = 1 To Me.costNeeds.Count
                    Dim costXML As clsOpenCostNeed = Me.costNeeds.Item(k - 1)

                    Dim dimension As Integer = costXML.monthlyNeeds.Length - 1
                    Dim newCost As New clsKostenart(anzahlMonate - 1)

                    If dimension = anzahlMonate - 1 And costXML.monthlyNeeds.Sum > 0 Then
                        ' alles in Ordnung , die Länge passt ...
                    Else
                        ' einfach die Summe hernehmen und verteilen ...
                        'ReDim oldWerte(0)
                        'oldWerte(0) = costXML.sum
                        ReDim Xwerte(anzahlMonate - 1)
                        Call .berechneBedarfe(Me.startDate.ToLocalTime, Me.finishDate.ToLocalTime, costXML.monthlyNeeds, 1.0, Xwerte)
                        costXML.monthlyNeeds = Xwerte
                    End If

                    If CostDefinitions.containsName(costXML.costName) Then
                        costUID = CInt(CostDefinitions.getCostdef(costXML.costName).UID)
                    Else
                        ' Kostenart existiert noch nicht
                        ' wird hier neu aufgenommen

                        Dim newCostDef As New clsKostenartDefinition
                        newCostDef.name = costXML.costName
                        newCostDef.farbe = RGB(120, 120, 120)

                        newCostDef.UID = CostDefinitions.Count + 1
                        If Not missingCostDefinitions.containsName(newCostDef.name) Then
                            missingCostDefinitions.Add(newCostDef)
                        End If

                        CostDefinitions.Add(newCostDef)

                        costUID = newCostDef.UID
                    End If


                    newCost.KostenTyp = costUID
                    newCost.Xwerte = costXML.monthlyNeeds

                    ' jetzt zur Phase dazu tun 
                    .AddCost(newCost)

                Next


                '.changeStartandDauer(phaseStartOffset, tmpDauer)

                '
                ' jetzt die Meilensteine aufnehmen ...
                '
                Try
                    Dim msAnzahl As Integer = Me.milestones.Count
                    For m = 1 To msAnzahl

                        Dim newresult As New clsMeilenstein(parent:=phase)

                        Try
                            Me.milestones.Item(m - 1).copyTo(newresult)
                            .addMilestone(newresult)
                        Catch ex As Exception

                        End Try

                    Next
                Catch ex As Exception

                End Try





            End With


        End Sub

        Sub New()

            name = ""
            originalName = ""
            abbreviation = ""
            appearance = ""

            color = 0
            breadCrumb = ""

            sourceUID = ""

            description = ""
            risks = New List(Of clsOpenRiskChance)

            startDate = Nothing
            finishDate = Nothing
            earliestStartOffset = 0
            latestStartOffset = 0

            responsible = ""

            ratings = New List(Of clsOpenRating)

            costNeeds = New List(Of clsOpenCostNeed)
            resourceNeeds = New List(Of clsOpenResourceNeed)

            milestones = New List(Of clsOpenMilestone)

        End Sub

    End Class
    ' ##################### Klasse clsOpenRiskChance
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsOpenRiskChance

        Public rcName As String
        Public description As String
        Public category As String
        Public potentialVariationInDuration As Integer
        Public damageBenefitValue As Double

        Private _probability As Double

        Public Property probability As Double
            Get
                probability = _probability
            End Get
            Set(value As Double)
                If value > 0 And value <= 1.0 Then
                    _probability = value
                Else
                    Throw New ArgumentException("Probability needs to be a value gretar than 0 and less or equal 1.0) ")
                End If
            End Set
        End Property


        ''' <summary>
        ''' kopiert von einer Hauptspeicher Struktur in die XML Struktur 
        ''' </summary>
        ''' <param name="riskChance"></param>
        ''' <remarks></remarks>
        Public Sub copyFrom(ByVal riskChance As clsRiskChance)

            With riskChance
                Me.rcName = .rcName
                Me.description = .description
                Me.category = .category
                Me.probability = .probability
                Me.potentialVariationInDuration = .potentialVariationInDuration
                Me.damageBenefitValue = .damageBenefitValue
            End With
        End Sub

        ''' <summary>
        ''' kopiert von einer XML Struktur in die Hauptspeicher Struktur 
        ''' </summary>
        ''' <param name="riskChance"></param>
        ''' <remarks></remarks>
        Public Sub copyTo(ByRef riskChance As clsRiskChance)

            With riskChance
                .rcName = Me.rcName
                .description = Me.description
                .category = Me.category
                .probability = Me.probability
                .potentialVariationInDuration = Me.potentialVariationInDuration
                .damageBenefitValue = Me.damageBenefitValue
            End With

        End Sub

        Sub New()

            rcName = "dummy"
            description = ""
            category = ""

            _probability = 0.5
            potentialVariationInDuration = 0
            damageBenefitValue = 0.0


        End Sub
    End Class
    ' ##################################### Klasse clsOpenCostNeed
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsOpenCostNeed

        Public costName As String
        Public sum As Double
        Public monthlyNeeds() As Double

        Sub copyFrom(ByVal cost As clsKostenart)

            With cost
                Me.costName = .name

                For i As Integer = 0 To .getDimension
                    Me.monthlyNeeds(i) = .Xwerte(i)
                Next

            End With

            ' ist zwar redundant, kann aber dann ggf in der Datenbank direkt abgefragt werden ... 
            Me.sum = Me.monthlyNeeds.Sum

        End Sub

        Sub New()
            ReDim monthlyNeeds(0)
            costName = ""
            sum = 0
        End Sub

        Sub New(ByVal dimension As Integer)
            ReDim monthlyNeeds(dimension)
            costName = ""
            sum = 0
        End Sub

    End Class
    ' ##################################### Klasse clsOpenResourceNeed
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsOpenResourceNeed

        Public resourceName As String
        Public sum As Double
        Public monthlyNeeds() As Double

        Sub copyFrom(ByVal role As clsRolle)

            With role
                Me.resourceName = .name

                For i As Integer = 0 To .getDimension
                    Me.monthlyNeeds(i) = .Xwerte(i)
                Next

            End With

            ' ist zwar redundant, kann aber dann ggf in der Datenbank direkt abgefragt werden ... 
            Me.sum = Me.monthlyNeeds.Sum

        End Sub

        Sub New()
            ReDim monthlyNeeds(0)
            resourceName = ""
            sum = 0
        End Sub

        Sub New(ByVal dimension As Integer)
            ReDim monthlyNeeds(dimension)
            resourceName = ""
            sum = 0
        End Sub
    End Class
    ' ##################################### Klasse clsOpenMilestone 
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsOpenMilestone

        Public name As String
        Public originalName As String
        Public abbreviation As String
        Public appearance As String
        Public color As Integer
        Public breadcrumb As String

        Public sourceUID As String

        Public description As String
        Public risks As List(Of clsOpenRiskChance)

        Public finishDate As Date
        Public earliestFinishOffset As Integer
        Public latestFinishOffset As Integer

        Public responsible As String

        Public ratings As List(Of clsOpenRating)
        Public deliverables As List(Of String)

        ''' <summary>
        ''' kopiert von einer Hauptspeicher Struktur in die XML Struktur
        ''' </summary>
        ''' <param name="hspMS"></param>
        ''' <remarks></remarks>
        Public Sub copyFrom(ByVal hspMS As clsMeilenstein)

            With hspMS
                Me.name = .name
                Me.originalName = .originalName
                Me.abbreviation = .shortName
                Me.appearance = ""
                Me.color = .farbe

                Me.breadcrumb = .Parent.parentProject.hierarchy.getBreadCrumb(.nameID)
                'Me.sourceUID = ""
                'Me.description = ""
                '' die Liste der Risiken ...

                Me.finishDate = .getDate.ToUniversalTime

                'Me.earliestFinishOffset = 0
                'Me.latestFinishOffset = 0

                Me.responsible = .verantwortlich

                ' jetzt die ratings kopieren 
                Try
                    For i = 1 To .bewertungsCount
                        Dim newRating As New clsOpenRating
                        newRating.copyfrom(.getBewertung(i))
                        ratings.Add(newRating)
                    Next
                Catch ex As Exception

                End Try

                ' jetzt die Deliverables kopieren ... 
                Try
                    For i = 1 To .countDeliverables
                        Dim deliv As String = .getDeliverable(i)
                        Me.deliverables.Add(deliv)
                    Next
                Catch ex As Exception

                End Try

            End With


        End Sub

        ''' <summary>
        ''' kopiert von einer XML Struktur in die Hauptspeicher Struktur 
        ''' </summary>
        ''' <param name="ms"></param>
        ''' <remarks></remarks>
        Public Sub copyTo(ByRef ms As clsMeilenstein)

            Dim i As Integer

            Try
                With ms


                    .nameID = .Parent.parentProject.hierarchy.findUniqueElemKey(Me.name, True)

                    .shortName = Me.abbreviation
                    .originalName = Me.originalName
                    .appearance = Me.appearance
                    .farbe = Me.color

                    .verantwortlich = Me.responsible
                    .offset = DateDiff(DateInterval.Day, .Parent.getStartDate, Me.finishDate.ToLocalTime)

                    ' die Deliverables übertragen 
                    For i = 1 To Me.deliverables.Count
                        Dim tmpDeliverable As String = Me.deliverables.Item(i - 1)
                        .addDeliverable(tmpDeliverable)
                    Next


                    For i = 1 To Me.ratings.Count

                        Dim newb As New clsBewertung
                        Try
                            Me.ratings.Item(i - 1).copyTo(newb)
                            .addBewertung(newb)
                        Catch ex1 As Exception

                        End Try

                    Next

                End With

            Catch ex As Exception

            End Try



        End Sub

        Sub New()
            name = ""
            originalName = ""
            abbreviation = ""
            appearance = ""
            color = 0
            breadcrumb = ""

            sourceUID = ""

            description = ""

            risks = New List(Of clsOpenRiskChance)

            finishDate = Nothing
            earliestFinishOffset = 0
            latestFinishOffset = 0

            responsible = ""

            ratings = New List(Of clsOpenRating)
            deliverables = New List(Of String)

        End Sub
    End Class

    Public Class clsOpenRating

        Public color As Integer
        Public description As String
        Public rater As String
        Public ratingDate As Date

        ''' <summary>
        ''' kopiert die Werte eine clsBewertung in clsOpenRating
        ''' </summary>
        ''' <param name="bewertung"></param>
        ''' <remarks></remarks>
        Public Sub copyFrom(ByVal bewertung As clsBewertung)

            With bewertung
                Me.color = .colorIndex
                Me.description = .description
                Me.rater = .bewerterName
                Me.ratingDate = .datum
            End With

        End Sub

        Public Sub copyTo(ByRef bewertung As clsBewertung)

            With bewertung
                .colorIndex = Me.color
                .description = Me.description
                .bewerterName = Me.rater
                .datum = Me.ratingDate
            End With

        End Sub

        Public Sub New()

            color = 0
            description = ""
            rater = ""
            ratingDate = Date.Now

        End Sub
    End Class
End Class
