Public Class clsLicences

    Private _allLicenceKeys As SortedList(Of String, String)

    ''' <summary>
    ''' errechnet aus einem maximalen Datum, einer User Kennung und einer Komponenten Kennung den Schlüssel 
    ''' </summary>
    ''' <param name="untilDate"></param>
    ''' <param name="User"></param>
    ''' <param name="komponente"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property berechneKey(ByVal untilDate As Date, ByVal User As String, ByVal komponente As String) As String
        Get
            Dim licKey As 
        End Get
    End Property

    Public Sub protokolliere(ByVal curDate As Date, ByVal user As String, ByVal komponente As String)

    End Sub

    Public Sub New()
        _allLicenceKeys = New SortedList(Of String, String)

    End Sub

End Class
