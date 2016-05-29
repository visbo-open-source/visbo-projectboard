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
    Public strategicFit As Double
    Public risk As Double

    ' die Custom Fields für ein Projekt 
    Public customDblFields As SortedList(Of Integer, Double)
    Public customStringFields As SortedList(Of Integer, String)
    Public customBoolFields As SortedList(Of Integer, Boolean)

    Public tasks As List(Of clsOpenTask)


    Public Sub copyFrom(ByVal projekt As clsProjekt)


        With projekt
            Me.projectName = calcProjektKey(projekt)
            Me.variantName = .variantName

            If Not IsNothing(.timeStamp) Then
                Me.timeStamp = .timeStamp.ToUniversalTime
            Else
                Me.timeStamp = Date.UtcNow
            End If

            ' die folgenden Infors werden noch nicht besetzt 
            Me.sourceDBURL = ""
            Me.sourceDBName = ""

            If Not IsNothing(.Id) Then
                Me.sourceUID = .Id
            End If

            Me.projectType = .VorlagenName

            Me.budget = .Erloes
            Me.currency = "€"

            Me.projectTitle = ""

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

            Next


        End With


    End Sub

    Public Sub copyTo(ByRef hproj As clsProjekt)

    End Sub

    ''' <summary>
    ''' Konstruktor für ein neues Projekt
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        projectName = "Testproject"
        variantName = ""
        timeStamp = Date.Now
        projectType = ""
        sourceDBURL = ""
        sourceDBName = ""
        sourceUID = ""

        projectStakeholder = New List(Of String)

        budget = 0.0
        currency = "€"

        projectTitle = ""
        strategicFit = 5
        risk = 5

        tasks = New List(Of clsOpenTask)

    End Sub
    ' ############################ Klasse clsOpenTask
    ''' <summary>
    ''' Klasse Phase
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsOpenTask

        Public categorizedName As String
        Public originalName As String
        Public abbreviation As String
        Public appearance As String
        Public color As Integer
        Public wbsCode As String

        Public sourceUID As String

        Public description As String
        Public risks As List(Of clsOpenRiskChance)

        Public startDate As Date
        Public finishDate As Date
        Public earliestStartOffset As Integer
        Public latestStartOffset As Integer

        Public responsible As String

        Public ratings As List(Of clsOpenRating)

        Public sCustomFields As SortedList(Of String, String)
        Public dCustomFields As SortedList(Of String, Double)
        Public bCustomFields As SortedList(Of String, Boolean)

        Public costNeeds As List(Of clsOpenCostNeed)
        Public resourceNeeds As List(Of clsOpenResourceNeed)

        Public milestones As List(Of clsOpenMilestone)

        Public Sub copyFrom(ByVal cPhase As clsPhase, _
                            Optional ByVal optResponsible As String = Nothing, _
                            Optional ByVal optDescription As String = Nothing)

            Dim r As Integer, k As Integer
            Dim dimension As Integer

            With cPhase
                Me.categorizedName = .name
                Me.originalName = .originalName
                Me.abbreviation = .shortName
                Me.appearance = .appearance
                Me.color = .farbe
                'Me.wbsCode = ""

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

                For r = 1 To .countMilestones
                    Dim newOpenMilestone As New clsOpenMilestone

                    Try
                        newOpenMilestone.copyFrom(.getMilestone(r))
                        milestones.Add(newOpenMilestone)
                    Catch ex As Exception

                    End Try

                Next

                For k = 1 To .countCosts
                    dimension = .getCost(k).getDimension
                    Dim newCost As New clsOpenCostNeed(dimension)
                    newCost.copyFrom(.getCost(k))
                    costNeeds.Add(newCost)
                Next

            End With


        End Sub

        Sub New()

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

        Public categorizedName As String
        Public originalName As String
        Public abbreviation As String
        Public appearance As String
        Public color As Integer
        Public wbsCode As String

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
                Me.categorizedName = .name
                Me.originalName = .originalName
                Me.abbreviation = .shortName
                Me.appearance = ""
                Me.color = .farbe

                'Me.wbsCode = ""
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

        End Sub

        Sub New()

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

        Public Sub New()

            color = 0
            description = ""
            rater = ""
            ratingDate = Date.Now

        End Sub
    End Class
End Class
