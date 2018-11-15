Public Class clsUserLoginSignup
    Public Property email As String
    Public Property password As String
    Public Property profile As clsUserProfile


    Sub New()
        _email = "not set"
        _password = "not set"
        _profile = New clsUserProfile
    End Sub
End Class
