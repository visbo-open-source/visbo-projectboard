Public Class clsConstellation

    Private allItems As Dictionary(Of String, clsConstellationItem)

    Public Property constellationName As String

    Public ReadOnly Property Liste() As Dictionary(Of String, clsConstellationItem)

        Get
            Liste = allItems
        End Get

    End Property
    Public ReadOnly Property getItem(index As Integer) As clsConstellationItem

        Get
            getItem = allItems(index)
        End Get

    End Property


    Public ReadOnly Property getItem(key As String) As clsConstellationItem

        Get
            getItem = allItems(key)
        End Get

    End Property

    Public ReadOnly Property count() As Integer

        Get
            count = allItems.Count
        End Get

    End Property

    Public Sub Add(cItem As clsConstellationItem)

        Dim key As String
        key = cItem.projectName & "#" & cItem.variantName
        allItems.Add(key, cItem)

    End Sub

    Public Sub Remove(index As Integer)

        allItems.Remove(index)

    End Sub

    Public Sub Remove(key As String)

        allItems.Remove(key)

    End Sub

    Sub New()

        allItems = New Dictionary(Of String, clsConstellationItem)

    End Sub

End Class
