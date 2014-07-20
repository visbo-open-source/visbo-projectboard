Public Class clsPhasen

    Private AllPhasen As Collection


    Public Sub Add(phase As clsPhasenDefinition)

        AllPhasen.Add(phase, phase.name)

    End Sub

    Public Sub Remove(myitem As Object)

        AllPhasen.Remove(myitem)

    End Sub

    Public ReadOnly Property Count() As Integer

        Get
            Count = AllPhasen.Count
        End Get

    End Property

    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = AllPhasen.Contains(name)
        End Get
    End Property

    Public ReadOnly Property getPhaseDef(ByVal myitem As String) As clsPhasenDefinition

        Get
            getPhaseDef = CType(AllPhasen.Item(myitem), clsPhasenDefinition)
        End Get

    End Property

    Public ReadOnly Property getPhaseDef(ByVal myitem As Integer) As clsPhasenDefinition

        Get
            getPhaseDef = CType(AllPhasen.Item(myitem), clsPhasenDefinition)
        End Get

    End Property

    Public Sub New()

        AllPhasen = New Collection
        
    End Sub

End Class
