Public Class clsRollen


    Private AllRollen As Collection


    Public Sub Add(roledef As clsRollenDefinition)

        Try
            AllRollen.Add(roledef, roledef.name)
        Catch ex As Exception
            Throw New ArgumentException(roledef.name & " existiert bereits")
        End Try


    End Sub

    Public Sub Remove(myitem As Object)

        AllRollen.Remove(myitem)

    End Sub
    '
    '
    '
    Public ReadOnly Property Count() As Integer

        Get

            Count = AllRollen.Count

        End Get

    End Property

    Public ReadOnly Property liste As Collection
        Get
            liste = AllRollen
        End Get
    End Property
    ''' <summary>
    ''' prüft ob name in der Collection enthalten ist
    ''' </summary>
    ''' <param name="name">Typ String</param>
    ''' <value></value>
    ''' <returns>wahr, wenn name enthalten ist; falsch, sonst</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = AllRollen.Contains(name)
        End Get
    End Property
    '
    '
    '
    Public ReadOnly Property getRoledef(ByVal myitem As String) As clsRollenDefinition

        Get

            Try
                getRoledef = AllRollen.Item(myitem)
            Catch ex As Exception
                Throw New ArgumentException(myitem & " gibt es nicht als Rolle")
            End Try

        End Get

    End Property

    Public ReadOnly Property getRoledef(ByVal myitem As Integer) As clsRollenDefinition

        Get

            Try
                getRoledef = AllRollen.Item(myitem)
            Catch ex As Exception
                Throw New ArgumentException(" es gibt keine Rolle mit Nummer " & myitem)
            End Try

        End Get

    End Property

    Public Sub New()

        AllRollen = New Collection

    End Sub

End Class
