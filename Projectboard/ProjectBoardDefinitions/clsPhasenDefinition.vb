Public Class clsPhasenDefinition

    Private uuid As Long            ' ur: 16.09.2019: wird aktuell nirgends berücksichtigt, auch keine Sortierung
    Private _farbe As Long
    Private _darstellungsKlasse As String
    Private Const defaultName As String = "Phasen Default"

    ' Name der Phase
    Public Property name As String

    ' Abkürzung, die in Reports für diese Phase verwendet werden soll 
    Public Property shortName As String


    'Public Property farbe As Object
    Public Property schwellWert As Integer

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
                'ur:190722
                '_arbe = appearanceDefinitions.Item(_darstellungsKlasse).form.Fill.ForeColor.RGB
                _farbe = appearanceDefinitions.liste.Item(_darstellungsKlasse).FGcolor
            Else

                _farbe = awinSettings.missingDefinitionColor

            End If

            farbe = _farbe

        End Get
        
    End Property


    ' Angabe der UID der Phase
    Public Property UID() As Long
        Get
            UID = uuid
        End Get
        Set(value As Long)
            uuid = value
        End Set
    End Property

    ''' <summary>
    ''' kopiert in Me die Werte der übergebenen Phasen-Definition
    ''' wenn der optionale Name angegeben ist, wird dieser Name, 
    ''' nicht der Name der übergebenen Phasen-Definition angegeben  
    ''' </summary>
    ''' <param name="phDef"></param>
    ''' <remarks></remarks>
    Public Sub copyFrom(ByVal phDef As clsPhasenDefinition, Optional ByVal newName As String = "")

        If Not IsNothing(phDef) Then
            With Me

                If newName = "" Then
                    .name = phDef.name
                Else
                    .name = newName
                End If
                .schwellWert = phDef.schwellWert
                .shortName = phDef.shortName
                .darstellungsKlasse = phDef.darstellungsKlasse
                '.farbe = phDef.farbe

            End With
        Else
            Throw New ArgumentException("Phase-Definition in Kopier-Funktion ist Nothing")
        End If
        

    End Sub

    Public Sub New()

        _name = ""
        _shortName = ""
        _darstellungsKlasse = defaultName
        _schwellWert = 0

        Try
            If appearanceDefinitions.liste.ContainsKey(_darstellungsKlasse) Then
                _farbe = appearanceDefinitions.liste.Item(_darstellungsKlasse).FGcolor
            End If
        Catch ex As Exception
            _farbe = CLng(RGB(120, 120, 120))
        End Try


    End Sub

End Class
