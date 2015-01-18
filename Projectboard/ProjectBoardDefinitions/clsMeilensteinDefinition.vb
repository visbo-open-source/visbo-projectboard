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
            If value = "" Or Not appearanceDefinitions.ContainsKey(value) Then
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
    ''' <value>setzt den Default Wert der Farbe für diese Phase, unabhängig von der Darstellungsklasse</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property farbe As Long
        Get
            Try

                If appearanceDefinitions.ContainsKey(_darstellungsKlasse) Then
                    _farbe = appearanceDefinitions.Item(_darstellungsKlasse).form.Fill.ForeColor.RGB
                End If

            Catch ex As Exception
                _farbe = appearanceDefinitions.Item(defaultName).form.Fill.ForeColor.RGB
            End Try

            farbe = _farbe

        End Get
        Set(value As Long)

            _farbe = value

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
