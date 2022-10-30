Public Class clsRankingParameters

    Private _projectName As String
    Public Property projectName As String
        Get
            projectName = _projectName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _projectName = value.Trim
            Else
                _projectName = "??"
            End If
        End Set
    End Property

    Private _peopleSuggestions As SortedList(Of String, Double)
    Public Property peopleSuggestions As SortedList(Of String, Double)
        Get
            peopleSuggestions = _peopleSuggestions
        End Get
        Set(value As SortedList(Of String, Double))
            If Not IsNothing(value) Then
                _peopleSuggestions = value
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
                _projectVariantName = value.Trim
            Else
                _projectVariantName = ""
            End If
        End Set
    End Property

    Private _newStartDate As Date
    Public Property newStartDate As Date
        Get
            newStartDate = _newStartDate
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _newStartDate = value
            Else
                _newStartDate = Date.MinValue
            End If
        End Set
    End Property

    Private _earliestStart As Date
    Public Property earliestStart As Date
        Get
            earliestStart = _earliestStart
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                If DateDiff(DateInterval.Day, Date.Now, value) > 0 Then
                    _earliestStart = value
                Else
                    _earliestStart = Date.Now.AddDays(1)
                End If

            Else
                _earliestStart = Date.Now.AddDays(1)
            End If
        End Set
    End Property

    Private _latestEnd As Date
    Public Property latestEnd As Date
        Get
            latestEnd = _latestEnd
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _latestEnd = value
            Else
                _latestEnd = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(13)
            End If
        End Set
    End Property

    Private _shortestDuration As Double
    Public Property shortestDuration As Double
        Get
            shortestDuration = _shortestDuration
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value > 0 Then
                    If value < 1.0 Then
                        _shortestDuration = value
                    ElseIf value > 5.0 Then
                        _shortestDuration = value
                    Else
                        _shortestDuration = 1.0
                    End If

                Else
                    _shortestDuration = 1.0
                End If
            Else
                _shortestDuration = 1.0
            End If
        End Set
    End Property

    Private _longestDuration As Double
    Public Property longestDuration As Double
        Get
            longestDuration = _longestDuration
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value > 1.0 Then
                    _longestDuration = value
                Else
                    _longestDuration = 1.0
                End If
            Else
                _longestDuration = 1.0
            End If
        End Set
    End Property

    Private _hedgeFactor As Double
    Public Property hedgeFactor As Double
        Get
            hedgeFactor = _hedgeFactor
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value < 1.0 Then
                    _hedgeFactor = value
                Else
                    _hedgeFactor = 1.0
                End If
            Else
                _hedgeFactor = 1.0
            End If
        End Set
    End Property



    Public Sub New()

        _peopleSuggestions = New SortedList(Of String, Double)

        _projectName = "ABC Test"
        _projectVariantName = ""
        _newStartDate = Date.MinValue
        _earliestStart = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(1)
        _latestEnd = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(13)
        _shortestDuration = 1.0
        _longestDuration = 1.0
        _hedgeFactor = 1.0

    End Sub

End Class
