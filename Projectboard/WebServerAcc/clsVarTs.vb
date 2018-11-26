Public Class clsVarTs
    ' VariantenNamen
    Public Property vname As String
    ' DatumUhrzeit, zudem der Cach - Short gefüllt wurde
    Public Property timeCShort As Date
    ' VPversions-Short sortiert nach timeStamp
    Public Property tsShort As SortedList(Of Date, clsProjektWebShort)
    ' DatumUhrzeit, zudem der Cach-Long gefüllt wurde
    Public Property timeCLong As Date
    ' VPversions-Long sortiert nach timeStamp
    Public Property tsLong As SortedList(Of Date, clsProjektWebLong)

    Sub New()
        _vname = ""
        _timeCShort = Date.MinValue
        _tsShort = New SortedList(Of Date, clsProjektWebShort)
        _timeCLong = Date.MinValue
        _tsLong = New SortedList(Of Date, clsProjektWebLong)
    End Sub
End Class
