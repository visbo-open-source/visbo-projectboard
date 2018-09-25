Public Class clsWebOneVC

    Public Property state As String
    Public Property message As String
    Public Property vc As clsVC

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vc = New clsVC
    End Sub

End Class
