Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsMeilensteine

    Private allMilestones As SortedList(Of String, clsMeilensteinDefinition)

    ''' <summary>
    ''' fügt der nach key=name sortierten Liste einen weiteren Eintrag hinzu 
    ''' </summary>
    ''' <param name="milestone"></param>
    ''' <remarks></remarks>
    Public Sub Add(milestone As clsMeilensteinDefinition)

        'Dim key As String = calcKey(milestone.name, milestone.belongsTo)


        If Not IsNothing(milestone) Then
            Dim key As String = milestone.name
            If allMilestones.ContainsKey(key) Then
                Throw New ArgumentException("Identifier " & key & _
                                            " existiert bereits!")
            Else
                allMilestones.Add(key, milestone)
            End If

        Else
            Throw New ArgumentException("Meilenstein Definition ist Nothing")
        End If
        

    End Sub


    ''' <summary>
    ''' gibt die Anzahl von Elementen in der Meilenstein Definition zurück  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count() As Integer

        Get
            Count = allMilestones.Count
        End Get

    End Property

    ''' <summary>
    ''' gibt die Meilenstein Definition an der Position "Index" zurück 
    ''' Nothing, wenn index kleiner Null oder größer Anzahl Elemente-^1
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property elementAt(ByVal index As Integer) As clsMeilensteinDefinition
        Get
            If index >= 0 And index <= Me.allMilestones.Count - 1 Then
                elementAt = Me.allMilestones.ElementAt(index).Value
            Else
                elementAt = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob der angegebene Meilenstein in der Liste vorkommt 
    ''' 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Contains(ByVal name As String) As Boolean

        Get
            'Dim key As String = calcKey(name, belongsTo)
            'Dim key As String = name

            Contains = allMilestones.ContainsKey(name)


        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl der Meilensteine mit Namen "name" zurück 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAnzahl(ByVal name As String) As Integer
        Get
            Dim anzahl As Integer = 0
            For Each ms As KeyValuePair(Of String, clsMeilensteinDefinition) In allMilestones
                If ms.Value.name = name Then
                    anzahl = anzahl + 1
                End If
            Next

            getAnzahl = anzahl

        End Get
    End Property

    ' ''' <summary>
    ' ''' gibt eine Collection von Elementen zurück, die alle den übergebenen Meilenstein als Meilenstein Namen haben  
    ' ''' </summary>
    ' ''' <param name="name"></param>
    ' ''' <value></value>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public ReadOnly Property getNameCollection(ByVal name As String) As Collection

    '    Get
    '        Dim tmpCollection As New Collection
    '        Dim key As String
    '        For Each ms As KeyValuePair(Of String, clsMeilensteinDefinition) In allMilestones
    '            If ms.Value.name = name Then
    '                key = calcKey(ms.Value.name, ms.Value.belongsTo)
    '                tmpCollection.Add(key)
    '            End If
    '        Next

    '        getNameCollection = tmpCollection

    '    End Get

    'End Property

    ''' <summary>
    ''' gibt den Meilenstein zurück, der den übergebenen Name hat 
    ''' 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneDef(ByVal name As String) As clsMeilensteinDefinition

        Get
            'Dim key As String = calcKey(name, belongsTo)

            If allMilestones.ContainsKey(name) Then
                getMilestoneDef = allMilestones.Item(name)
            Else
                getMilestoneDef = Nothing
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt den Meilenstein an Position Index zurück
    ''' Index muss Zahl zwischen 1 und Anzahl Elemente sein
    ''' wenn Zahl ausserhalb liegt, wird die leere Menge zurückgegegen 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneDef(ByVal index As Integer) As clsMeilensteinDefinition
        Get

            If index > 0 And index <= allMilestones.Count Then
                getMilestoneDef = allMilestones.ElementAt(index - 1).Value
            Else
                getMilestoneDef = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die Shape Definition für den angegebenen Meilenstein zurück 
    ''' wenn es den nicht gibt, wird die Default Milestone Klasse verwendet 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShape(ByVal name As String) As xlNS.Shape
        Get
            Dim appearanceID As String
            Dim defaultMilestoneAppearance As String = "Meilenstein Default"

            'Dim key As String = calcKey(name, belongsTo)

            If allMilestones.ContainsKey(name) Then
                appearanceID = allMilestones.Item(name).darstellungsKlasse
                If appearanceID = "" Then
                    appearanceID = defaultMilestoneAppearance
                End If
            Else
                appearanceID = defaultMilestoneAppearance
            End If

            ' jetzt ist in der AppearanceID was drin ... 
            getShape = appearanceDefinitions.Item(appearanceID).form

        End Get
    End Property

    ''' <summary>
    ''' gibt die Abkürzung, den Shortname für den Meilenstein zurück
    ''' wenn er nicht gefunden wird: "-"
    ''' </summary>
    ''' <param name="name">Langname Meilenstein</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAbbrev(ByVal name As String) As String
        Get
            Dim msAbbrev As String = "-"

            'Dim key As String = calcKey(name, belongsTo)

            If allMilestones.ContainsKey(name) Then
                msAbbrev = allMilestones.Item(name).shortName
            End If

            getAbbrev = msAbbrev

        End Get
    End Property

    Public Sub New()
        allMilestones = New SortedList(Of String, clsMeilensteinDefinition)
    End Sub

    ''' <summary>
    ''' löscht die Meilenstein-Definition mit dem übergebenen Namen aus der Liste , sofern vorhanden
    ''' wenn nicht vorhanden, keine Änderung; aber auch keine Mitteilung 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub remove(ByVal name As String)

        If allMilestones.ContainsKey(name) Then
            allMilestones.Remove(name)
        End If

    End Sub

    Public Sub Clear()
        allMilestones.Clear()
    End Sub

End Class
