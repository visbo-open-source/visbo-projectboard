Public Class clsFilterDefinitions

    Public filterListe As SortedList(Of String, clsFilter)
    Private currentFilter As String
    Private isActive As Boolean


    ''' <summary>
    ''' gibt den Filter mit Namen name zurück
    ''' wenn er nicht existiert, dann Nothing
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property retrieveFilter(ByVal name As String) As clsFilter
        Get
            If filterListe.ContainsKey(name) Then
                retrieveFilter = filterListe.Item(name)
            Else
                retrieveFilter = Nothing
            End If
        End Get
    End Property


    ''' <summary>
    ''' speichert einen Filter unter dem angegebenen Namen in den Filter Definitionen 
    ''' wenn der Name schon existiert, wird der Filter entsprechend überschrieben 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub storeFilter(ByVal name As String, ByVal filter As clsFilter)


        If filterListe.ContainsKey(name) Then
            Dim ok As Boolean = filterListe.Remove(name)
        End If

        Call filterListe.Add(name, filter)


    End Sub
    Public ReadOnly Property Liste As SortedList(Of String, clsFilter)

        Get
            Liste = filterListe
        End Get

    End Property
    Sub New()
        filterListe = New SortedList(Of String, clsFilter)
    End Sub
End Class
