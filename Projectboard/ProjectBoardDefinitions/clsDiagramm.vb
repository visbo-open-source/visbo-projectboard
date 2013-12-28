Public Class clsDiagramm

    Public Property diagrammTyp As String
    Public Property DiagrammTitel As String
    Public Property isCockpitChart As Boolean
    Public Property top As Double
    Public Property left As Double
    Public Property width As Double
    Public Property height As Double
    Public Property kennung As String


    'Private eventvariable As clsAwinEvent
    Private eventvariable As Object
    Private myCollection As Collection

    Public WriteOnly Property setDiagramEvent() As Object

        'Set(awinevent As clsAwinEvent)
        '    eventvariable.ChartEvents = awinevent.ChartEvents
        'End Set

        Set(awinevent As Object)
            eventvariable = awinevent
        End Set

    End Property

    'Public WriteOnly Property setpfDiagramEvent() As clsAwinEvent

    '    Set(awinevent As clsAwinEvent)
    '        eventvariable.pfChartEvents = awinevent.pfChartEvents
    '    End Set

    'End Property

    'gs war vorher getCollection bzw setCollection; muss entsprechend ausgetauscht werden
    Public Property gsCollection() As Collection
        Get
            gsCollection = myCollection
        End Get

        Set(awinCollection As Collection)
            myCollection = awinCollection
        End Set
    End Property

    Public Sub New()
        eventvariable = New Object
        myCollection = New Collection
        _top = 0.0
        _left = 0.0
        _width = 0.0
        _height = 0.0
        _kennung = ""
    End Sub


End Class
