Public Class clsRollenDefinition

    Public Property name As String
    Public Property farbe As Object
    Public Property Startkapa As Double
    Public Property tagessatzIntern As Double
    Public Property tagessatzExtern As Double
    Public Property kapazitaet As Double()

    Private uuid As Long
    Private Kapa() As Double

    Public Property UID() As Long

        Get

            UID = uuid

        End Get

        Set(value As Long)

            uuid = value

        End Set

    End Property

    Public Sub New()

        ReDim kapazitaet(120)

    End Sub

End Class
