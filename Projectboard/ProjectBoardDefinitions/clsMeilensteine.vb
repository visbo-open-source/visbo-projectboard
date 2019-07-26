Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsMeilensteine

    Private _allMilestones As SortedList(Of String, clsMeilensteinDefinition)

    ''' <summary>
    ''' fügt der nach key=name sortierten Liste einen weiteren Eintrag hinzu 
    ''' </summary>
    ''' <param name="milestone"></param>
    ''' <remarks></remarks>
    Public Sub Add(milestone As clsMeilensteinDefinition)

        'Dim key As String = calcKey(milestone.name, milestone.belongsTo)


        If Not IsNothing(milestone) Then
            Dim key As String = milestone.name
            If _allMilestones.ContainsKey(key) Then
                ' einfach nichts machen 
                'Throw New ArgumentException("Identifier " & key & _
                '                            " existiert bereits!")
            Else
                _allMilestones.Add(key, milestone)
            End If

        Else
            ' nichts machen
            'Throw New ArgumentException("Meilenstein Definition ist Nothing")
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
            Count = _allMilestones.Count
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
            If index >= 0 And index <= Me._allMilestones.Count - 1 Then
                elementAt = Me._allMilestones.ElementAt(index).Value
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

            Contains = _allMilestones.ContainsKey(name)


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
            For Each ms As KeyValuePair(Of String, clsMeilensteinDefinition) In _allMilestones
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

            If _allMilestones.ContainsKey(name) Then
                getMilestoneDef = _allMilestones.Item(name)
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

            If index > 0 And index <= _allMilestones.Count Then
                getMilestoneDef = _allMilestones.ElementAt(index - 1).Value
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

            If _allMilestones.ContainsKey(name) Then
                appearanceID = _allMilestones.Item(name).darstellungsKlasse
                If appearanceID = "" Then
                    appearanceID = defaultMilestoneAppearance
                End If
            Else
                appearanceID = defaultMilestoneAppearance
            End If

            '' ''Dim ok As Boolean = False
            '' ''While Not ok
            '' ''    Try
            '' ''        ' jetzt ist in der AppearanceID was drin ... 
            '' ''        getShape = appearanceDefinitions.Item(appearanceID).form
            '' ''        If Not IsNothing(getShape) Then
            '' ''            ok = True
            '' ''        Else
            '' ''            Call MsgBox("nothing")
            '' ''        End If
            '' ''    Catch ex As Exception
            '' ''        Call MsgBox("getshape fehlerhaft")
            '' ''        getShape = Nothing
            '' ''    End Try

            '' ''End While


            ' jetzt ist in der AppearanceID was drin ... 
            getShape = appearanceDefinitions.Item(appearanceID).form

            'ur:190725
            'Dim appear As clsAppearance = appearanceDefinitions.Item(appearanceID)
            'getShape = CType(appInstance.Worksheets(arrWsNames(ptTables.MPT)),
            '    Microsoft.Office.Interop.Excel.Worksheet).Shapes.AddShape(appear.shpType, 0, 0, appear.width, appear.height)

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
            Dim msAbbrev As String = name

            If _allMilestones.ContainsKey(name) Then
                msAbbrev = _allMilestones.Item(name).shortName
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
            If _allMilestones.ContainsKey(name) Then
                tmpErg = _allMilestones.Item(name).darstellungsKlasse
            End If
            getAppearance = tmpErg
        End Get
    End Property


    Public Sub New()
        _allMilestones = New SortedList(Of String, clsMeilensteinDefinition)
    End Sub

    ''' <summary>
    ''' löscht die Meilenstein-Definition mit dem übergebenen Namen aus der Liste , sofern vorhanden
    ''' wenn nicht vorhanden, keine Änderung; aber auch keine Mitteilung 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub remove(ByVal name As String)

        If _allMilestones.ContainsKey(name) Then
            _allMilestones.Remove(name)
        End If

    End Sub

    Public Sub Clear()
        _allMilestones.Clear()
    End Sub


    ''' <summary>
    ''' gibt die komplette Liste zurück
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property liste() As SortedList(Of String, clsMeilensteinDefinition)
        Get
            liste = _allMilestones
        End Get
    End Property


End Class
