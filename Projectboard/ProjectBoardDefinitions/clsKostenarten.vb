Public Class clsKostenarten

    Private AllKostenarten As Collection


    Public Sub Add(costdef As clsKostenartDefinition)

        Try
            AllKostenarten.Add(costdef, costdef.name)
        Catch ex As Exception
            Throw New ArgumentException(costdef.name & " existiert bereits")
        End Try


    End Sub

    Public Sub Remove(myitem As Object)

        Try
            AllKostenarten.Remove(myitem)
        Catch ex As Exception
            Throw New ArgumentException(myitem & " kann nicht als Kostenart entfernt werden")
        End Try


    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Count = AllKostenarten.Count
        End Get
    End Property

    ''' <summary>
    ''' prüft, ob name in der Kostenarten Collection enthalten ist 
    ''' </summary>
    ''' <param name="name">typ string</param>
    ''' <value></value>
    ''' <returns>wahr, wenn enthalten; falsch sonst</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = AllKostenarten.Contains(name)
        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As String) As clsKostenartDefinition
        Get

            Try
                getCostdef = AllKostenarten.Item(myitem)
            Catch ex As Exception
                Throw New ArgumentException(myitem & " ist keine Kostenart")
            End Try

        End Get
    End Property

    Public ReadOnly Property getCostdef(ByVal myitem As Integer) As clsKostenartDefinition
        Get
            Try
                getCostdef = AllKostenarten.Item(myitem)
            Catch ex As Exception
                Throw New ArgumentException(" es gibt keine Kostenart mit Nummer " & myitem)
            End Try

        End Get
    End Property


    Public Sub New()
        AllKostenarten = New Collection
    End Sub

End Class
