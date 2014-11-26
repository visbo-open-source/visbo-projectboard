Public Class clsPhasenDefinition

    Private uuid As Long

    ' Name der Phase
    Public Property name As String


    Public Property farbe As Object
    Public Property schwellWert As Integer
    Public Property darstellungsKlasse As String



    ' Angabe der UID der Phase
    Public Property UID() As Long
        Get
            UID = uuid
        End Get
        Set(value As Long)
            uuid = value
        End Set
    End Property

    Public Sub New()
        _name = ""
        _darstellungsKlasse = ""
        _schwellWert = 0
        _farbe = RGB(120, 120, 120)
    End Sub

End Class
