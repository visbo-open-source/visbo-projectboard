Public Class clsConstellationDB

    Public allItems As List(Of clsConstellationItemDB)
    Public constellationName As String
    Public Id As String

    Sub copyfrom(ByRef c As clsConstellation)

        Me.constellationName = c.constellationName
        
        For Each item In c.Liste
            Dim newItem As New clsConstellationItemDB
            newItem.copyfrom(item.Value)
            Me.allItems.Add(newItem)
        Next

    End Sub

    Sub copyto(ByRef c As clsConstellation)
        Dim key As String

        c.constellationName = Me.constellationName

        For Each item In Me.allItems
            Dim newItem As New clsConstellationItem
            item.copyto(newItem)
            key = item.projectName & "#" & item.variantName
            c.Liste.Add(key, newItem)
        Next


    End Sub

    Public Class clsConstellationItemDB
        Public projectName As String
        Public variantName As String
        Public Start As Date
        Public show As Boolean
        Public zeile As Integer

        Sub copyfrom(ByRef item As clsConstellationItem)

            With item
                Me.projectName = .projectName
                Me.variantName = .variantName
                Me.Start = .Start.ToUniversalTime
                Me.show = .show
                Me.zeile = .zeile
            End With
        End Sub

        Sub copyto(ByRef item As clsConstellationItem)

            With item
                .projectName = Me.projectName
                .variantName = Me.variantName
                .Start = Me.Start.ToLocalTime
                .show = Me.show
                .zeile = Me.zeile
            End With

        End Sub

        Sub New()

        End Sub

    End Class

    Sub New()
        allItems = New List(Of clsConstellationItemDB)
    End Sub

End Class
