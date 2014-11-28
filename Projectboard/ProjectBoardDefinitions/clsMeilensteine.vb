Imports xlNS = Microsoft.Office.Interop.Excel
Public Class clsMeilensteine

    Private allMilestones As SortedList(Of String, clsMeilensteinDefinition)

    ''' <summary>
    ''' fügt der nach key=name+belongsTO sortierten Liste einen weiteren Eintrag hinzu 
    ''' </summary>
    ''' <param name="milestone"></param>
    ''' <remarks></remarks>
    Public Sub Add(milestone As clsMeilensteinDefinition)

        Dim key As String = calcKey(milestone.name, milestone.belongsTo)

        If allMilestones.ContainsKey(key) Then
            Throw New ArgumentException("Identifier " & milestone.UID.ToString & _
                                        " existiert bereits!")
        Else
            allMilestones.Add(key, milestone)
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
    ''' gibt zurück, ob der angegebene Meilenstein in der Liste vorkommt 
    ''' 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Contains(ByVal name As String, ByVal belongsTo As String) As Boolean

        Get
            Dim key As String = calcKey(name, belongsTo)

            Contains = allMilestones.ContainsKey(key)


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

    ''' <summary>
    ''' gibt eine Collection von Elementen zurück, die alle den übergebenen Meilenstein als Meilenstein Namen haben  
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getNameCollection(ByVal name As String) As Collection

        Get
            Dim tmpCollection As New Collection
            Dim key As String
            For Each ms As KeyValuePair(Of String, clsMeilensteinDefinition) In allMilestones
                If ms.Value.name = name Then
                    key = calcKey(ms.Value.name, ms.Value.belongsTo)
                    tmpCollection.Add(key)
                End If
            Next

            getNameCollection = tmpCollection

        End Get

    End Property

    ''' <summary>
    ''' gibt den Meilenstein zurück, der den übergebenen Name und Kennzeichnung "belongsTo" hat 
    ''' 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="belongsTo"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneDef(ByVal name As String, ByVal belongsTo As String) As clsMeilensteinDefinition

        Get
            Dim key As String = calcKey(name, belongsTo)

            If allMilestones.ContainsKey(key) Then
                getMilestoneDef = allMilestones.Item(key)
            Else
                getMilestoneDef = Nothing
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die Shape Definition für den angegebenen Meilenstein zurück 
    ''' wenn es die Kombination name, belongsto nicht gibt, wird nur Name als Suchkriterium verwendet 
    ''' wenn es auch den nicht gibt, wird die Default Milestone Klasse verwendet 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="belongsTo"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShape(ByVal name As String, ByVal belongsTo As String) As xlNS.Shape
        Get
            Dim appearanceID As String
            Dim tmpCollection As Collection
            Dim defaultMilestoneAppearance As String = "Meilenstein Default"

            Dim key As String = calcKey(name, belongsTo)

            If allMilestones.ContainsKey(key) Then
                appearanceID = allMilestones.Item(key).darstellungsKlasse
                If appearanceID = "" Then
                    appearanceID = defaultMilestoneAppearance
                End If
            Else

                tmpCollection = Me.getNameCollection(name)
                If tmpCollection.Count = 0 Then
                    appearanceID = defaultMilestoneAppearance
                Else
                    appearanceID = allMilestones.Item(CStr(tmpCollection.Item(1))).darstellungsKlasse
                    If appearanceID = "" Then
                        appearanceID = defaultMilestoneAppearance
                    End If
                End If
            End If

            ' jetzt ist in der AppearanceID was drin ... 
            getShape = appearanceDefinitions.Item(appearanceID).form

        End Get
    End Property

    ''' <summary>
    ''' gibt die Abkürzung, den Shortname für den Meilenstein zurück
    ''' wenn er nicht gefunden wird: "n.a."
    ''' </summary>
    ''' <param name="name">Langname Meilenstein</param>
    ''' <param name="belongsTo">Phasen-Name (nur wichtig, wenn Meilenstein Namen mehrfach vorkommen</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAbbrev(ByVal name As String, ByVal belongsTo As String) As String
        Get
            Dim msAbbrev As String = "n.a."

            Dim key As String = calcKey(name, belongsTo)

            If allMilestones.ContainsKey(key) Then
                msAbbrev = allMilestones.Item(key).shortName
            End If

            getAbbrev = msAbbrev

        End Get
    End Property

    ''' <summary>
    ''' gibt die Abkürzung zu einem gegebenen Meilenstein zurück; eine Phase muss nicht angegegen werden; er sucht und findet das erste Vorkommen
    ''' diese Vorgehensweise liefert nur korrekte Ergebnisse, wenn sichergestellt ist, daß keine Duplikate in den Namen vorkommen 
    ''' </summary>
    ''' <param name="msName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAbbrev(ByVal msName As String) As String
        Get
            Dim msAbbrev As String = "n.a."
            Dim i As Integer = 0
            Dim anzahl As Integer = allMilestones.Count - 1
            Dim found As Boolean = False
            Dim msDefinition As clsMeilensteinDefinition

            While i <= anzahl And Not found
                msDefinition = allMilestones.ElementAt(i).Value
                If msDefinition.name = msName Then
                    found = True
                    msAbbrev = msDefinition.shortName
                End If
                i = i + 1
            End While

            getAbbrev = msAbbrev

        End Get
    End Property

    Public Sub New()
        allMilestones = New SortedList(Of String, clsMeilensteinDefinition)
    End Sub

    Private Function calcKey(ByVal name As String, ByVal belongsTo As String) As String


        If IsNothing(belongsTo) Then
            belongsTo = ""
        End If

        If belongsTo = "" Then
            calcKey = name
        Else
            calcKey = belongsTo & "#" & name
        End If


    End Function



End Class
