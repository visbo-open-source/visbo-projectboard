Public Class clsBusinessUnit

    Private _name As String
    Private _color As Long

    Public Property name As String
        Get
            name = _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public Property color As Long
        Get
            color = _color
        End Get
        Set(value As Long)
            If value >= 0 Then
                _color = value
            Else
                Throw New ArgumentException("negativer Wert für Farbe nicht zugelassen ...")
            End If
        End Set
    End Property

    Sub New()
        _name = ""
        _color = 0
    End Sub

End Class
