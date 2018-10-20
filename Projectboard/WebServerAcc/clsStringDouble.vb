Public Class clsStringDouble
    Public Property str As String
    Public Property dbl As Double

    Sub New()
        _str = ""
        _dbl = 0.0
    End Sub
    Sub New(ByVal str As String, ByVal dbl As Double)
        _str = str
        _dbl = dbl
    End Sub
End Class
