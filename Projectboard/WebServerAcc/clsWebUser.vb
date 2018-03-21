Public Class clsWebUser

    Public Property state As String
    Public Property message As String
    Public Property user As clsUser

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _user = New clsUser
    End Sub

End Class
