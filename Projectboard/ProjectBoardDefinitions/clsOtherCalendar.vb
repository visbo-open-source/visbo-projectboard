Public Class clsOtherCalendar

    Public Property otherCal As SortedList(Of String, clsFirstWDLastWD)
    Sub New()
        _otherCal = New SortedList(Of String, clsFirstWDLastWD)
    End Sub

End Class
Public Class clsFirstWDLastWD
    Public Property firstWorkDay As Date
    Public Property lastWorkDay As Date
    Sub New()
        firstWorkDay = Date.MinValue
        lastWorkDay = Date.MaxValue
    End Sub
End Class
