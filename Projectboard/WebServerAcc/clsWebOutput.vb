Public Class clsWebOutput
    Public Property state As String
    Public Property message As String

    Sub New()
        _state = "failure"
        _message = "not yet any response"
    End Sub
End Class
