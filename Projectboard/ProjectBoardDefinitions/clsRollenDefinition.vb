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

        ' Änderung 29.5.14 damit man zwanzig Jahre vom Start der Projekt-Tafel betrachten kann 
        ' Kapazität: die Null Position hat keine Bedeutung; kapazität(1) = der Wert für StartofCalendar
        ReDim kapazitaet(240)

    End Sub

End Class
