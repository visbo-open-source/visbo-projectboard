Public Class clsDiagramme
    'Private AllDiagrams As Collection
    Private AllDiagrams As SortedList(Of String, clsDiagramm)

    Public Sub New()
        'AllDiagrams = New Collection
        AllDiagrams = New SortedList(Of String, clsDiagramm)
    End Sub

    Public Sub Add(diagram As clsDiagramm)


        Try

            If AllDiagrams.ContainsKey(diagram.kennung) Then
                AllDiagrams.Remove(diagram.kennung)
            End If

            AllDiagrams.Add(diagram.kennung, diagram)
        Catch ex As Exception

        End Try


    End Sub

    Public Sub Clear()
        AllDiagrams.Clear()
    End Sub

    Public Sub Remove(myitem As Integer)

        AllDiagrams.RemoveAt(myitem - 1)
        'AllDiagrams.Remove(myitem)

    End Sub

    Public Sub Remove(kennung As String)

        Try
            AllDiagrams.Remove(kennung)
        Catch ex As Exception

        End Try


        'Dim i As Integer
        'Dim anzahl As Integer = AllDiagrams.Count
        'Dim found As Boolean
        'Dim hobj As clsDiagramm

        'i = 1
        'found = False
        'While i <= anzahl And Not found
        '    hobj = AllDiagrams.Item(i)
        '    If hobj.DiagrammTitel = title Then
        '        found = True
        '    Else
        '        i = i + 1
        '    End If
        'End While

        'If found Then
        '    AllDiagrams.Remove(i)
        'Else
        '    Throw New ArgumentException("Objekt nicht gefunden")
        'End If

    End Sub

    'Public Sub Remove(title As String, isCockpitChart As Boolean)
    '    Dim i As Integer
    '    Dim anzahl As Integer = AllDiagrams.Count
    '    Dim found As Boolean
    '    Dim hobj As clsDiagramm

    '    i = 1
    '    found = False
    '    While i <= anzahl And Not found
    '        hobj = AllDiagrams.ElementAt(i).Value
    '        If hobj.DiagrammTitel = title And hobj.isCockpitChart = isCockpitChart Then
    '            found = True
    '        Else
    '            i = i + 1
    '        End If
    '    End While

    '    If found Then
    '        AllDiagrams.Remove(i)
    '    Else
    '        Throw New ArgumentException("Objekt nicht gefunden")
    '    End If

    'End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Count = AllDiagrams.Count
        End Get
    End Property

    Public ReadOnly Property getDiagramm(myitem As Integer) As clsDiagramm
        Get
            'getDiagramm = AllDiagrams.Item(myitem)
            ' sotedlist elementat ist zero-based, im Gegensatz zur Collection
            getDiagramm = AllDiagrams.ElementAt(myitem - 1).Value
        End Get
    End Property

    Public ReadOnly Property getDiagramm(myitem As String) As clsDiagramm

        Get
            If AllDiagrams.ContainsKey(myitem) Then
                getDiagramm = AllDiagrams(myitem)
            Else
                getDiagramm = Nothing
            End If
        End Get

    End Property

    ''' <summary>
    ''' gibt zurück, ob die Diagramm-Liste ein Diagramm mit Schlüssel key enthält
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property contains(ByVal key As String) As Boolean
        Get
            contains = AllDiagrams.ContainsKey(key)
        End Get
    End Property

End Class
