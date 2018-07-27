Public Class clsWebVPf

    Public Property state As String
    Public Property message As String
    Public Property vpf As List(Of clsVPf)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vpf = New List(Of clsVPf)
    End Sub
End Class
