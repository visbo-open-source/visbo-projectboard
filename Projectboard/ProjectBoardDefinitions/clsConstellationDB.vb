Public Class clsConstellationDB

    Public allItems As List(Of clsConstellationItemDB)
    Public constellationName As String
    Public Id As String

    Sub copyfrom(ByVal c As clsConstellation)

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
            If Not c.Liste.ContainsKey(key) Then
                c.Liste.Add(key, newItem)
            Else
                Dim a As Integer = 0
                'Call MsgBox("Fehler bei Aufbau Konstellation mit Elem: " & key)
            End If

        Next


    End Sub

    Public Class clsConstellationItemDB
        Public projectName As String = ""
        Public variantName As String = ""
        Public Start As Date = StartofCalendar.AddMonths(-1)
        Public show As Boolean = True
        Public zeile As Integer = 0
        ' warum wird die entsprechende Projektvariante aufgenommen 
        Public reasonToInclude As String = ""
        ' warum wird die entsprechende Projektvariante nicht aufgenommen 
        Public reasonToExclude As String = ""

        Sub copyfrom(ByVal item As clsConstellationItem)

            With item
                Me.projectName = .projectName
                Me.variantName = .variantName
                Me.Start = .start.ToUniversalTime
                Me.show = .show
                Me.zeile = .zeile
                Me.reasonToInclude = .reasonToInclude
                Me.reasonToExclude = .reasonToExclude
            End With
        End Sub

        Sub copyto(ByRef item As clsConstellationItem)

            With item
                .projectName = Me.projectName
                .variantName = Me.variantName
                .start = Me.Start.ToLocalTime
                .show = Me.show
                .zeile = Me.zeile
                .reasonToInclude = Me.reasonToInclude
                .reasonToExclude = Me.reasonToExclude
            End With

        End Sub

        Sub New()
            
        End Sub

    End Class

    Sub New()
        allItems = New List(Of clsConstellationItemDB)
    End Sub

End Class
