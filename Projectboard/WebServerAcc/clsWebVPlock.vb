Public Class clsWebVPlock

    Public Property state As String
    Public Property message As String
    Public Property lock As List(Of clsVPLock)
    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _lock = New List(Of clsVPLock)
    End Sub
End Class
