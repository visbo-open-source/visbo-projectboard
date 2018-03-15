Public Class clsWebAllVC

    Public Property state As String
    Public Property message As String
    Public Property vc As List(Of clsVC)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vc = New List(Of clsVC)
    End Sub
End Class
