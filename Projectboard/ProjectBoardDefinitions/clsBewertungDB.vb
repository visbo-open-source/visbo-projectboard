Public Class clsBewertungDB
    ' Änderung tk: 2.11 deliverables / Ergebnisse hinzugefügt 

    Public color As Integer
    Public description As String
    Public deliverables As String
    Public bewerterName As String
    Public datum As Date

    Friend Sub CopyTo(ByRef newB As clsBewertung)

        With newB
            .colorIndex = Me.color
            .description = Me.description
            '.deliverables = Me.deliverables
            .datum = Me.datum
            .bewerterName = Me.bewerterName
        End With

    End Sub

    Friend Sub Copyfrom(ByVal b As clsBewertung)

        Me.color = b.colorIndex
        Me.description = b.description
        'Me.deliverables = b.deliverables
        Me.bewerterName = b.bewerterName
        Me.datum = b.datum

    End Sub

    Sub New()
        bewerterName = ""
        datum = Nothing
        color = 0
        description = ""
        deliverables = ""
    End Sub

End Class
