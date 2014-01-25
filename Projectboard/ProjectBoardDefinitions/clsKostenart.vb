Public Class clsKostenart

    Private typus As Integer
    Private Bedarf() As Double

    Public Property KostenTyp() As Integer
        Get
            KostenTyp = typus
        End Get

        Set(value As Integer)
            typus = value
        End Set

    End Property

    Public ReadOnly Property getDimension As Integer
        Get
            getDimension = Xwerte.Length - 1
        End Get
    End Property

    Public Property Xwerte() As Double()
        Get
            Xwerte = Bedarf
        End Get

        Set(values As Double())

            Bedarf = values

        End Set

    End Property

    Public Property Xwerte(ByVal index As Integer) As Double
        Get
            Xwerte = Bedarf(index)
        End Get

        Set(value As Double)
            Bedarf(index) = value
        End Set

    End Property

    Public ReadOnly Property name() As String
        Get
            name = CostDefinitions.getCostdef(typus).name
        End Get
    End Property

    Public ReadOnly Property farbe() As Object
        Get
            farbe = CostDefinitions.getCostdef(typus).farbe
        End Get
    End Property

    Public ReadOnly Property summe() As Double
        Get
            Dim isum As Double
            Dim i As Integer
            Dim ende As Integer

            ende = UBound(Bedarf)
            isum = 0
            For i = 0 To ende
                isum = isum + Bedarf(i)
            Next i

            summe = isum
        End Get
    End Property

    Public Sub CopyTo(ByRef newcost As clsKostenart)

        With newcost
            .KostenTyp = typus
            .Xwerte = Bedarf

        End With

    End Sub

    Public Sub New()

    End Sub

    Public Sub New(ByVal laenge As Integer)

        ReDim Bedarf(laenge)

    End Sub

End Class
