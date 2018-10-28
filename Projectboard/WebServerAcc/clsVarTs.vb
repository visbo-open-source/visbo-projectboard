Public Class clsVarTs
    Public Property vname As String
    Public Property timeCached As Date
    Public Property tsShort As SortedList(Of Date, clsProjektWebShort)
    Public Property tsLong As SortedList(Of Date, clsProjektWebLong)

    Sub New()
        _vname = ""
        _timeCached = Date.MinValue
        _tsShort = New SortedList(Of Date, clsProjektWebShort)
        _tsLong = New SortedList(Of Date, clsProjektWebLong)
    End Sub
End Class
