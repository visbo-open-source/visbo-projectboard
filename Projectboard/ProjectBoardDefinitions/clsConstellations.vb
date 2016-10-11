Public Class clsConstellations
    Private allConstellations As SortedList(Of String, clsConstellation)

    Public ReadOnly Property Count As Integer

        Get
            Count = allConstellations.Count
        End Get

    End Property

    Public ReadOnly Property Liste As SortedList(Of String, clsConstellation)

        Get
            Liste = allConstellations
        End Get

    End Property

    Public ReadOnly Property getConstellation(name As String) As clsConstellation
        Get

            If allConstellations.ContainsKey(name) Then
                getConstellation = allConstellations.Item(name)
            Else
                getConstellation = Nothing
            End If

        End Get
    End Property

    

    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = allConstellations.ContainsKey(name)
        End Get
    End Property

    Sub Add(ByVal item As clsConstellation)

        Try
            allConstellations.Add(item.constellationName, item)
        Catch ex As Exception
            Throw New ArgumentException("Konstellations-Name existiert bereits")
        End Try


    End Sub

    Sub Remove(ByVal key As String)

        Try
            allConstellations.Remove(key)
        Catch ex As Exception
            Throw New ArgumentException("Konstellation" & " key " & "konnte nicht gelöscht werden ")
        End Try

    End Sub

    Sub New()

        allConstellations = New SortedList(Of String, clsConstellation)

    End Sub

End Class
