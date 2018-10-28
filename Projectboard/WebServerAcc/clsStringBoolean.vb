Public Class clsStringBoolean
    Public Property str As String
    Public Property bool As Boolean

    Sub New()
        _str = ""
        _bool = False
    End Sub
    Sub New(ByVal str As String, ByVal bool As Boolean)
        _str = str
        _bool = bool
    End Sub

End Class
