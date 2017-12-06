Public Class clsPPTTimeMachine

    'Private currentTSIndex As Integer = -1
    Private _timeStamps As SortedList(Of Date, Boolean)
    Private _timeStampsIndex As Integer
    Private _anzahlShapesOnSlide As Integer

    Public Property timeStamps As SortedList(Of Date, Boolean)
        Get
            timeStamps = _timeStamps
        End Get
        Set(value As SortedList(Of Date, Boolean))
            If Not IsNothing(value) Then
                _timeStamps = value
            Else
                _timeStamps = Nothing
            End If
        End Set
    End Property

    Public Property timeStampsIndex As Integer
        Get
            timeStampsIndex = _timeStampsIndex
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If IsNumeric(value) Then
                    _timeStampsIndex = value
                Else
                    _timeStampsIndex = -1
                End If
            Else
                _timeStampsIndex = -1
            End If
        End Set
    End Property

    Public Property anzahlShapesOnSlide As Integer
        Get
            anzahlShapesOnSlide = _anzahlShapesOnSlide
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If IsNumeric(value) Then
                    _anzahlShapesOnSlide = value
                Else
                    _anzahlShapesOnSlide = 0
                End If
            Else
                _anzahlShapesOnSlide = 0
            End If
        End Set
    End Property



    Private Enum ptNavigationButtons
        letzter = 0
        erster = 1
        nachher = 2
        vorher = 3
        individual = 4
    End Enum


    Public Sub New()
        timeStamps = New SortedList(Of Date, Boolean)
        timeStampsIndex = -1
        anzahlShapesOnSlide = currentSlide.Shapes.Count
    End Sub
End Class
