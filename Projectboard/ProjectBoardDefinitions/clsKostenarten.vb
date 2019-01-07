Public Class clsKostenarten

    'Private _allKostenarten As Collection
    Private _allKostenarten As SortedList(Of Integer, clsKostenartDefinition)


    Public Sub Add(costdef As clsKostenartDefinition)

        If Not IsNothing(costdef) Then
            If Not _allKostenarten.ContainsKey(costdef.UID) Then
                _allKostenarten.Add(costdef.UID, costdef)
            Else
                Throw New ArgumentException(costdef.UID.ToString & " existiert bereits")
            End If
        Else
            Throw New ArgumentException("Kostenart darf nicht Nothing sein")
        End If


        ''Try
        ''    _allKostenarten.Add(Item:=costdef, Key:=costdef.name)
        ''Catch ex As Exception
        ''    Throw New ArgumentException(costdef.name & " existiert bereits")
        ''End Try


    End Sub

    Public ReadOnly Property liste() As SortedList(Of Integer, clsKostenartDefinition)
        Get
            liste = _allKostenarten
        End Get
    End Property

    ''Public Sub Remove(myitem As Object)

    ''    Try
    ''        _allKostenarten.Remove(myitem)
    ''    Catch ex As Exception
    ''        Throw New ArgumentException("Fehler bei Kostenart entfernen")
    ''    End Try


    ''End Sub


    ''' <summary>
    ''' liefert true zurück, wenn alle Kostendefinitionen der einen Liste identisch mit der anderen sind
    ''' </summary>
    ''' <param name="vglDefinitionen"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglDefinitionen As clsKostenarten)
        Get
            Dim stillIdentical = True

            If Me.Count = vglDefinitionen.Count Then
                Dim i As Integer = 0
                Do While i < _allKostenarten.Count And stillIdentical
                    stillIdentical = _allKostenarten.ElementAt(i).Value.isIdenticalTo(vglDefinitionen.getCostdef(i + 1))
                    i = i + 1
                Loop

            Else
                stillIdentical = False
            End If

            isIdenticalTo = stillIdentical
        End Get
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Count = _allKostenarten.Count
        End Get
    End Property

    ''' <summary>
    ''' prüft, ob name in der Kostenarten Collection enthalten ist 
    ''' </summary>
    ''' <param name="name">typ string</param>
    ''' <value></value>
    ''' <returns>wahr, wenn enthalten; falsch sonst</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsName(name As String) As Boolean
        Get

            Dim found As Boolean = False
            If IsNothing(name) Then
                ' found bleibt auf false
            Else
                Dim ix As Integer = 0
                Do While ix <= _allKostenarten.Count - 1 And Not found
                    If _allKostenarten.ElementAt(ix).Value.name = name Then
                        found = True
                    Else
                        ix = ix + 1
                    End If
                Loop
            End If

            containsName = found

        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob der Key bereits enthalten ist 
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsUid(uid As Integer) As Boolean
        Get

            containsUid = _allKostenarten.ContainsKey(uid)

        End Get
    End Property

    ''' <summary>
    ''' gibt die Kostenart mit ID = uid zurück; Nothing, wenn sie nicht existiert
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCostDefByID(ByVal uid As Integer) As clsKostenartDefinition
        Get
            If _allKostenarten.ContainsKey(uid) Then
                getCostDefByID = _allKostenarten.Item(uid)
            Else
                getCostDefByID = Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As String) As clsKostenartDefinition
        Get

            Dim tmpValue As clsKostenartDefinition = Nothing

            Dim found As Boolean = False
            Dim ix As Integer = 0

            Do While ix <= _allKostenarten.Count - 1 And Not found
                If _allKostenarten.ElementAt(ix).Value.name = myitem Then
                    found = True
                    tmpValue = _allKostenarten.ElementAt(ix).Value
                Else
                    ix = ix + 1
                End If
            Loop

            getCostdef = tmpValue

        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As Integer) As clsKostenartDefinition
        Get


            If myitem > 0 And myitem <= _allKostenarten.Count Then
                getCostdef = _allKostenarten.ElementAt(myitem - 1).Value
            Else
                getCostdef = Nothing
            End If


        End Get
    End Property


    Public Sub New()
        _allKostenarten = New SortedList(Of Integer, clsKostenartDefinition)
    End Sub

End Class
