Imports ProjectBoardDefinitions
Public Class clsJobParameters

    Friend kennung As PTRpa


    Private _allowedOverloadMonth As Double
    Public Property allowedOverloadMonth As Double
        Get
            allowedOverloadMonth = _allowedOverloadMonth
        End Get
        Set(value As Double)
            If value > 0.1 Then
                _allowedOverloadMonth = value
            End If
        End Set
    End Property

    Private _allowedOverloadTotal As Double
    Public Property allowedOverloadTotal As Double
        Get
            allowedOverloadTotal = _allowedOverloadTotal
        End Get
        Set(value As Double)
            If value > 0.1 Then
                _allowedOverloadTotal = value
            End If
        End Set
    End Property

    Private _limitPhases As Integer
    Public Property limitPhases As Integer
        Get
            limitPhases = _limitPhases
        End Get
        Set(value As Integer)
            If value > 0 Then
                _limitPhases = value
            End If
        End Set
    End Property

    Private _limitMilestones As Integer
    Public Property limitMilestones As Integer
        Get
            limitMilestones = _limitMilestones
        End Get
        Set(value As Integer)
            If value > 0 Then
                _limitMilestones = value
            End If
        End Set
    End Property

    Private _phases As Collection
    Public Property phases As Collection
        Get
            phases = _phases
        End Get
        Set(value As Collection)
            If Not IsNothing(value) Then
                _phases = value
            End If
        End Set
    End Property

    Public ReadOnly Property getPhaseNames() As List(Of String)
        Get
            Dim result As New List(Of String)
            For Each phName As String In _phases
                result.Add(phName)
            Next
            getPhaseNames = result
        End Get
    End Property

    Public Sub AddPhase(ByVal myPhase As String)
        If myPhase <> "" Then
            If Not _phases.Contains(myPhase) Then
                _phases.Add(myPhase, myPhase)
            End If
        End If
    End Sub

    Private _milestones As Collection
    Public Property milestones As Collection
        Get
            milestones = _milestones
        End Get
        Set(value As Collection)
            If Not IsNothing(value) Then
                _milestones = value
            End If
        End Set
    End Property

    Private _roleNames As Collection
    Public Property roleNames As Collection
        Get
            roleNames = _roleNames
        End Get
        Set(value As Collection)
            If Not IsNothing(value) Then
                _roleNames = value
            End If
        End Set
    End Property

    Private _costNames As Collection
    Public Property costNames As Collection
        Get
            costNames = _costNames
        End Get
        Set(value As Collection)
            If Not IsNothing(value) Then
                _costNames = value
            End If
        End Set
    End Property

    Private _revenueTitle As String
    Public Property revenueTitle As String
        Get
            revenueTitle = _revenueTitle
        End Get
        Set(value As String)
            If value <> "" Then
                _revenueTitle = value
            End If
        End Set
    End Property

    Public ReadOnly Property getMilestoneNames() As List(Of String)
        Get
            Dim result As New List(Of String)
            For Each msName As String In _milestones
                result.Add(msName)
            Next
            getMilestoneNames = result
        End Get
    End Property
    Public Sub AddMilestone(ByVal myMilestone As String)
        If myMilestone <> "" Then
            If Not _milestones.Contains(myMilestone) Then
                _milestones.Add(myMilestone, myMilestone)
            End If
        End If
    End Sub

    Private _considerRoleSkills As Collection
    Public Property considerRoleSkills As Collection
        Get
            considerRoleSkills = _considerRoleSkills
        End Get
        Set(value As Collection)
            If Not IsNothing(value) Then
                _considerRoleSkills = value
            End If
        End Set
    End Property
    Public Sub addRoleSkill(ByVal myRSName As String)
        If myRSName <> "" Then
            If RoleDefinitions.containsName(myRSName) Then
                If Not _considerRoleSkills.Contains(myRSName) Then
                    _considerRoleSkills.Add(myRSName, myRSName)
                End If
            End If
        End If
    End Sub

    Private _donotConsiderRoleSkills As Collection

    Public Property donotConsiderRoleSkills As Collection
        Get
            donotConsiderRoleSkills = _donotConsiderRoleSkills
        End Get
        Set(value As Collection)
            If Not IsNothing(value) Then
                _donotConsiderRoleSkills = value
            End If
        End Set
    End Property

    Public Sub minusRoleSkill(ByVal myRSName As String)
        If myRSName <> "" Then
            If RoleDefinitions.containsName(myRSName) Then
                If Not _donotConsiderRoleSkills.Contains(myRSName) Then
                    _donotConsiderRoleSkills.Add(myRSName, myRSName)
                End If
            End If
        End If
    End Sub

    Private _portfolioName As String
    Public Property portfolioName As String
        Get
            portfolioName = _portfolioName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _portfolioName = value
            End If
        End Set
    End Property

    Private _portfolioVariantName As String
    Public Property portfolioVariantName As String
        Get
            portfolioVariantName = _portfolioVariantName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _portfolioVariantName = value
            End If
        End Set
    End Property

    Private _projectVariantName As String
    Public Property projectVariantName As String
        Get
            projectVariantName = _projectVariantName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _projectVariantName = value
            End If
        End Set
    End Property

    Private _defaultLatestEnd As Date
    Public Property defaultLatestEnd As Date
        Get
            defaultLatestEnd = _defaultLatestEnd
        End Get
        Set(value As Date)
            If value > StartofCalendar Then
                _defaultLatestEnd = value
            End If
        End Set
    End Property

    Private _defaultDeltaInDays As Integer
    Public Property defaultDeltaInDays As Integer
        Get
            defaultDeltaInDays = _defaultDeltaInDays
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If value > 0 Then
                    _defaultDeltaInDays = value
                End If
            End If
        End Set
    End Property

    Private _changeFactorResourceNeeds As Double
    Public Property changeFactorResourceNeeds As Double
        Get
            changeFactorResourceNeeds = _changeFactorResourceNeeds
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value >= 0 Then
                    _changeFactorResourceNeeds = value
                End If
            End If
        End Set
    End Property

    Private _changeFactorDuration As Double
    Public Property changeFactorDuration As Double
        Get
            changeFactorDuration = _changeFactorDuration
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value >= 0 Then
                    _changeFactorDuration = value
                End If
            End If
        End Set
    End Property


    Private _sortItem As String
    Public Property sortItem As String
        Get
            sortItem = _sortItem
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _sortItem = value
            End If
        End Set
    End Property

    Private _templateName As String
    Public Property templateName As String
        Get
            templateName = _templateName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _templateName = value
            End If
        End Set
    End Property

    Private _compareWithFirstBaseline As Boolean
    Public Property compareWithFirstBaseline As Boolean
        Get
            compareWithFirstBaseline = _compareWithFirstBaseline
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _compareWithFirstBaseline = value
            End If
        End Set
    End Property
    Sub New()

        _sortItem = ""
        _allowedOverloadMonth = 1.0
        _allowedOverloadTotal = 1.0
        _limitPhases = 1
        _limitMilestones = 1
        _defaultDeltaInDays = 7
        _milestones = New Collection
        _phases = New Collection
        _roleNames = New Collection
        _costNames = New Collection
        _revenueTitle = "Revenue/Benefit"
        kennung = PTRpa.visboUnknown

        _projectVariantName = "Var1"
        _portfolioName = ""
        _portfolioVariantName = "Var1"
        _defaultLatestEnd = DateSerial(Date.Now.Year + 1, 12, 31)
        _templateName = "TMS"
        _compareWithFirstBaseline = False
        _changeFactorResourceNeeds = 1.0
        _changeFactorDuration = 1.0

    End Sub

End Class
