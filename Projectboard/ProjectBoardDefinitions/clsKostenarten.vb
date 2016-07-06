Public Class clsKostenarten

    Private _allKostenarten As Collection


    Public Sub Add(costdef As clsKostenartDefinition)

        Try
            _allKostenarten.Add(Item:=costdef, Key:=costdef.name)
        Catch ex As Exception
            Throw New ArgumentException(costdef.name & " existiert bereits")
        End Try


    End Sub

    Public Sub Remove(myitem As Object)

        Try
            _allKostenarten.Remove(myitem)
        Catch ex As Exception
            Throw New ArgumentException("Fehler bei Kostenart entfernen")
        End Try


    End Sub

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
            containsName = _allKostenarten.Contains(name)
        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As String) As clsKostenartDefinition
        Get

            Try
                getCostdef = CType(_allKostenarten.Item(myitem), clsKostenartDefinition)
            Catch ex As Exception
                Throw New ArgumentException(myitem & " ist keine Kostenart")
            End Try

        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As Integer) As clsKostenartDefinition
        Get
            Try
                getCostdef = CType(_allKostenarten.Item(myitem), clsKostenartDefinition)
            Catch ex As Exception
                Throw New ArgumentException(" es gibt keine Kostenart mit Nummer " & myitem)
            End Try

        End Get
    End Property


    Public Sub New()
        _allKostenarten = New Collection
    End Sub

End Class
