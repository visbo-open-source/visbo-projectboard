Public Class clsKostenart

    Private _typus As Integer
    Private _bedarf() As Double

    Public Property KostenTyp() As Integer
        Get
            KostenTyp = _typus
        End Get

        Set(value As Integer)
            _typus = value
        End Set

    End Property

    Public ReadOnly Property getDimension As Integer
        Get
            getDimension = _bedarf.Length - 1
        End Get
    End Property


    Public Property Xwerte() As Double()
        Get
            Xwerte = _bedarf
        End Get

        Set(values As Double())

            Dim ub As Integer = UBound(values)
            Dim tmpArray() As Double
            ReDim tmpArray(ub)

            For i As Integer = 0 To ub
                tmpArray(i) = values(i)
            Next
            _bedarf = tmpArray

        End Set

    End Property

    Public Property Xwerte(ByVal index As Integer) As Double
        Get
            Xwerte = _bedarf(index)
        End Get

        Set(value As Double)
            _bedarf(index) = value
        End Set

    End Property

    Public ReadOnly Property name() As String
        Get
            name = CostDefinitions.getCostdef(_typus).name
        End Get
    End Property

    Public ReadOnly Property farbe() As Object
        Get
            farbe = CostDefinitions.getCostdef(_typus).farbe
        End Get
    End Property

    Public ReadOnly Property summe() As Double
        Get
            Dim isum As Double
            Dim i As Integer
            Dim ende As Integer

            ende = UBound(_bedarf)
            isum = 0
            For i = 0 To ende
                isum = isum + _bedarf(i)
            Next i

            summe = isum
        End Get
    End Property

    Public Sub CopyTo(ByRef newcost As clsKostenart)

        With newcost
            .KostenTyp = _typus
            .Xwerte = _bedarf

        End With

    End Sub

    Public Sub New()

    End Sub

    Public Sub New(ByVal laenge As Integer)

        ReDim _bedarf(laenge)

    End Sub

End Class
