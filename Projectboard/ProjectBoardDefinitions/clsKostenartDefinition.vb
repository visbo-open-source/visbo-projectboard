Public Class clsKostenartDefinition

    Public name As String
    Public farbe As Object

    Private Budget() As Double
    Private uuid As Long

    Public Property UID() As Long
        Get
            UID = uuid
        End Get
        Set(value As Long)
            uuid = value
        End Set
    End Property

    Public Sub New()

    End Sub
End Class
