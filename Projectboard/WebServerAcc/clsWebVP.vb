Public Class clsWebVP

    Public Property state As String
    Public Property message As String
    Public Property vp As List(Of clsVP)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vp = New List(Of clsVP)
    End Sub
End Class
