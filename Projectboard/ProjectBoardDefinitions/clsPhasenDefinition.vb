Public Class clsPhasenDefinition

    Public Property name As String
    Public Property farbe As Object
    Public Property schwellWert As Integer

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
