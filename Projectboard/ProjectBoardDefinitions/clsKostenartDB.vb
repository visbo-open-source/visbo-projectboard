Public Class clsKostenartDB

    Public KostenTyp As Integer
    Public name As String
    Public farbe As Object
    Public Bedarf() As Double

    Sub copyFrom(ByVal cost As clsKostenart)

        With cost
            Me.KostenTyp = .KostenTyp
            Me.name = .name
            Me.farbe = .farbe
            Bedarf = .Xwerte
        End With

    End Sub

    Sub copyto(ByRef cost As clsKostenart)

        With cost
            .KostenTyp = Me.KostenTyp
            .Xwerte = Bedarf
        End With

    End Sub

    Sub New()

    End Sub

    Sub New(ByVal laenge As Integer)

        ReDim Bedarf(laenge)

    End Sub


End Class
