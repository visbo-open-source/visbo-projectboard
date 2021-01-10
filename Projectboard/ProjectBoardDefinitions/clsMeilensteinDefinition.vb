Public Class clsMeilensteinDefinition

    Private _darstellungsKlasse As String
    Private Const defaultName As String = "Meilenstein Default"
    Private _farbe As Long

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
            If value = "" Or Not appearanceDefinitions.liste.ContainsKey(value) Then
                _darstellungsKlasse = defaultName
            Else
                _darstellungsKlasse = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest die Farbe entsprechend der Definition der Darstellungsklasse 
    ''' wenn es die nicht gibt, wird der Default für diese Phase verwendet  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property farbe As Long
        Get

            If appearanceDefinitions.liste.ContainsKey(_darstellungsKlasse) Then
                'ur: 19022
                'farbe = appearanceDefinitions.Item(_darstellungsKlasse).form.Fill.ForeColor.RGB
                _farbe = appearanceDefinitions.liste.Item(_darstellungsKlasse).FGcolor
            Else
                _farbe = awinSettings.AmpelNichtBewertet
            End If


            farbe = _farbe

        End Get

    End Property

    ' Angabe der UID des Meilensteins
    Public Property UID As Long

    ''' <summary>
    ''' kopiert in Me die Werte der übergebenen Phasen-Definition
    ''' wenn der optionale Name angegeben ist, wird dieser Name, 
    ''' nicht der Name der übergebenen Phasen-Definition angegeben  
    ''' </summary>
    ''' <param name="msDef"></param>
    ''' <remarks></remarks>
    Public Sub copyFrom(ByVal msDef As clsMeilensteinDefinition, Optional ByVal newName As String = "")

        If Not IsNothing(msDef) Then
            With Me

                If newName = "" Then
                    .name = msDef.name
                Else
                    .name = newName
                End If
                .schwellWert = msDef.schwellWert
                .shortName = msDef.shortName
                .darstellungsKlasse = msDef.darstellungsKlasse
                '.farbe = msDef.farbe
                .belongsTo = msDef.belongsTo

            End With
        Else
            Throw New ArgumentException("Phase-Definition in Kopier-Funktion ist Nothing")
        End If


    End Sub

    Public Sub New()
        _name = ""
        _shortName = ""
        _belongsTo = ""
        _schwellWert = 0
        _darstellungsKlasse = defaultName
        _UID = -1
    End Sub

End Class
