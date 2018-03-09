Public Class clsAllVC
    Public Property state As String
    Public Property message As String
    Public Property vc As List(Of clsVC)
    Sub New()
        _state = "failure"
        _message = "not yet response"
        _vc = New List(Of clsVC)
    End Sub
End Class
