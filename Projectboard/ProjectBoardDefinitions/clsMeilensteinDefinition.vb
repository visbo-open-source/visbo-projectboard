Public Class clsMeilensteinDefinition

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
    Public Property darstellungsKlasse As String

    ' Angabe der UID des Meilensteins
    Public Property UID As Long

    Public Sub New()
        _name = ""
        _shortName = ""
        _belongsTo = ""
        _schwellWert = 0
        _darstellungsKlasse = ""
        _UID = -1
    End Sub

End Class
