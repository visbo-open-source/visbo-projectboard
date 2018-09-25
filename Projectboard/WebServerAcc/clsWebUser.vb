Public Class clsWebUser

    Public Property state As String
    Public Property message As String
    Public Property user As clsUserReg

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _user = New clsUserReg
    End Sub

End Class
