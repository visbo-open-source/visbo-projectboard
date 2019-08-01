Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsPhasen

    Private _allPhasen As SortedList(Of String, clsPhasenDefinition)


    ''' <summary>
    ''' nimmt die Phase auf; wenn der Name bereits vergeben ist, wird nichts gemacht ...
    ''' wenn PhaseDef = Nothing, wird auch nichts gemacht 
    ''' es werden keine Exceptions geworfen; wenn man an der Aufruf Stelle wissen muss, ob der Name vergeben ist, muss über .contains geprüft werden 
    ''' </summary>
    ''' <param name="phaseDef"></param>
    ''' <remarks></remarks>
    Public Sub Add(phaseDef As clsPhasenDefinition)

        If Not IsNothing(phaseDef) Then
            If Not _allPhasen.ContainsKey(phaseDef.name) Then
                _allPhasen.Add(phaseDef.name, phaseDef)
            Else
                ' nichts tun , ist ja schon da 
            End If
        Else
            ' nichts tun , es ist ja nichts aufzunehmen  
        End If



    End Sub

    Public ReadOnly Property Count() As Integer

        Get
            Count = _allPhasen.Count
        End Get

    End Property


    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = _allPhasen.ContainsKey(name)
        End Get
    End Property

    Public ReadOnly Property getPhaseDef(ByVal myitem As String) As clsPhasenDefinition

        Get
            If _allPhasen.ContainsKey(myitem) Then
                getPhaseDef = CType(_allPhasen.Item(myitem), clsPhasenDefinition)
            Else
                'getPhaseDef = AllPhasen.First.Value
                getPhaseDef = Nothing
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die Phasen-Definition an der Index-Position index zurüclk: Index kann von 1 .. Anzahl Phasedefs gehen 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseDef(ByVal index As Integer) As clsPhasenDefinition

        Get
            If index < 1 Then
                index = 1
            ElseIf index > _allPhasen.Count Then
                index = _allPhasen.Count
            End If
            getPhaseDef = CType(_allPhasen.ElementAt(index - 1).Value, clsPhasenDefinition)
        End Get

    End Property

    ''' <summary>
    ''' gibt die Abkürzung, den Shortname für den Meilenstein zurück
    ''' wenn er nicht gefunden wird: 
    ''' </summary>
    ''' <param name="name">Langname Phase</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAbbrev(ByVal name As String) As String
        Get
            Dim msAbbrev As String = name

            'Dim key As String = calcKey(name, belongsTo)

            If _allPhasen.ContainsKey(name) Then
                msAbbrev = CType(_allPhasen.Item(name), clsPhasenDefinition).shortName
            End If

            getAbbrev = msAbbrev

        End Get
    End Property

    ''' <summary>
    ''' gibt die Darstellungsklasse des Elements zurück 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAppearance(ByVal name As String) As String
        Get
            Dim tmpErg As String = ""
            If _allPhasen.ContainsKey(name) Then
                tmpErg = _allPhasen.Item(name).darstellungsKlasse
            End If
            getAppearance = tmpErg
        End Get
    End Property

    '''' <summary>
    '''' gibt die Shape Definition für die angegebene Phase zurück
    '''' wenn es die Definition für name nicht gibt, wird die Default Phasen Klasse verwendet   
    '''' </summary>
    '''' <param name="name"></param>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public ReadOnly Property getShape(ByVal name As String) As xlNS.Shape
    '    Get
    '        Dim appearanceID As String
    '        Dim defaultPhaseAppearance As String = "Phasen Default"


    '        If _allPhasen.ContainsKey(name) Then

    '            appearanceID = CType(_allPhasen.Item(name), clsPhasenDefinition).darstellungsKlasse
    '            If appearanceID = "" Then
    '                appearanceID = defaultPhaseAppearance
    '            End If

    '        Else

    '            appearanceID = defaultPhaseAppearance

    '        End If

    '        getShape = appearanceDefinitions.Item(appearanceID).form

    '    End Get
    'End Property

    ''' <summary>
    ''' gibt die Shape Appearance Definition für die angegebene Phase zurück
    ''' wenn es die Definition für name nicht gibt, wird die Default Phasen Klasse verwendet   
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapeApp(ByVal name As String) As clsAppearance
        Get
            Dim appearanceID As String
            Dim defaultPhaseAppearance As String = "Phasen Default"


            If _allPhasen.ContainsKey(name) Then

                appearanceID = CType(_allPhasen.Item(name), clsPhasenDefinition).darstellungsKlasse
                If appearanceID = "" Then
                    appearanceID = defaultPhaseAppearance
                End If

            Else

                appearanceID = defaultPhaseAppearance

            End If

            ' jetzt ist in der AppearanceID was drin ...

            getShapeApp = appearanceDefinitions.Item(appearanceID)

        End Get
    End Property

    ''' <summary>
    ''' löscht die Phasen-Definition mit dem übergebenen Namen aus der Liste , sofern vorhanden
    ''' wenn nicht vorhanden, keine Änderung; aber auch keine Mitteilung 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub remove(ByVal name As String)

        If _allPhasen.ContainsKey(name) Then
            _allPhasen.Remove(name)
        End If

    End Sub

    ''' <summary>
    ''' leert die komplette Liste 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()

        _allPhasen.Clear()

    End Sub

    ''' <summary>
    ''' gibt die komplette Liste zurück
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property liste() As SortedList(Of String, clsPhasenDefinition)
        Get
            liste = _allPhasen
        End Get
    End Property

    Public Sub New()

        _allPhasen = New SortedList(Of String, clsPhasenDefinition)


    End Sub

End Class
