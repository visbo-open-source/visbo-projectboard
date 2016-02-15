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
            berechneKey = ""
        End Get
    End Property

    ''' <summary>
    ''' checkt, ob ein gültiger Lizez-KEy vorhanden ist
    ''' dazu werden alle ausgelesenen Lizenzkeys mit den Eingabe Werten user, komponente verglichen 
    ''' Es sollten folgende Meldungen kommen: 
    '''  
    ''' </summary>
    ''' <param name="user"></param>
    ''' <param name="komponente"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property validLicence(ByVal user As String, ByVal komponente As String) As Boolean
        Get
            Dim heute As Date = Date.Now
            validLicence = True

        End Get
    End Property


    Public Sub New()
        _allLicenceKeys = New SortedList(Of String, String)

    End Sub

End Class
