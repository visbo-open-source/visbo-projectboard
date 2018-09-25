Public Class clsWebLongVPv

    Public Property state As String
    Public Property message As String

    Public Property vpv As List(Of clsProjektWebLong)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vpv = New List(Of clsProjektWebLong)
    End Sub
End Class
