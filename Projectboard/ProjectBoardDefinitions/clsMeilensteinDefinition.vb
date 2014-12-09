Public Class clsMeilensteinDefinition

    Private _darstellungsKlasse As String
    Private Const defaultName As String = "Meilenstein Default"

    ' Name des Meilensteine
    Public Property name As String

    ' Abkürzung, die in Reports für diesen Meilenstein verwendet werden soll 
    Public Property shortName As String

    ' ID/Name der Phase, zu der der Meilenstein gehört; wenn Null, dann zum Projekt 
    'Public Property belongsTo As Long; nach Re-Factoring Phasen-Klassen muss das auf die ID verweisen 
    Public Property belongsTo As String

    ' Angabe eines optionalen Schwellwerts; kann verwendet werden für Leistbarkeitsanalysen 
    Public Property schwellWert As Integer

    ' Angabe der Darstellungsklasse
    ''' <summary>
    ''' liest schreibt die Darstellungsklasse; 
    ''' beim Schreiben wird der Name durch den Default Namen ersetzt, wenn er nicht in den Darstellungsklassen auftaucht  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property darstellungsKlasse As String
        Get
            darstellungsKlasse = _darstellungsKlasse
        End Get

        Set(value As String)
            If value = "" Or Not appearanceDefinitions.ContainsKey(value) Then
                _darstellungsKlasse = defaultName
            Else
                _darstellungsKlasse = value
            End If
        End Set
    End Property

    ' Angabe der UID des Meilensteins
    Public Property UID As Long

    Public Sub New()
        _name = ""
        _shortName = ""
        _belongsTo = ""
        _schwellWert = 0
        _darstellungsKlasse = defaultName
        _UID = -1
    End Sub

End Class
