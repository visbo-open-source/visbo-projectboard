Public Class clsAllVC
    Public Property state As String
    Public Property message As String
    Public Property vc() As clsVC
    Sub New()
        _state = "failure"
        _message = "not yet response"
        _vc = New clsVC()
    End Sub
End Class
