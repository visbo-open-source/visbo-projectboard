Public Class clsDefaultCalendar
    ' sortierte Liste auf relMonth ausgehend von startOfCal
    Public Property defCal As SortedList(Of Integer, clsBusinessDays)
    Public Sub New()
        _defCal = New SortedList(Of Integer, clsBusinessDays)
    End Sub
End Class
Public Class clsBusinessDays

    Public Property year As Integer
    Public Property month As Integer
    Public Property noOfBusinessDays As Integer
    Public Property noOfNonBusinessDays As Integer
    Public Sub New()
        _year = DateAndTime.Year(Date.MinValue)
        _month = DateAndTime.Month(Date.MinValue)
        _noOfBusinessDays = 0
        _noOfNonBusinessDays = 0
    End Sub

End Class
