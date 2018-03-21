Public Class clsInputSignupLogin
    Public Property name As String
    Public Property email As String
    Public Property password As String
    Public Property phone As String
    Public Property company As String
    Sub New()
        _name = "unknown"
        _email = "max.mustermann@t-online.de"
        _password = "xxxxxx"
        _phone = ""
        _company = "unknown"
    End Sub
End Class
