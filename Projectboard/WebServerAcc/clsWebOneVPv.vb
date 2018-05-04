Public Class clsWebOneVPv

    Public Property state As String
    Public Property message As String

    Public Property vpv As List(Of clsProjektWeblong)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vpv = New List(Of clsProjektWeblong)
    End Sub
End Class
