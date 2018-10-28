Public Class clsStringString
    Public Property strkey As String
    Public Property strvalue As String

    Sub New()
        _strkey = ""
        _strvalue = ""
    End Sub
    Sub New(ByVal strkey As String, ByVal strvalue As String)
        _strkey = strkey
        _strvalue = strvalue
    End Sub
End Class
