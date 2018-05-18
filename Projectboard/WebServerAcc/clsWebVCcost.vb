Public Class clsWebVCcost

    Public Property state As String
    Public Property message As String
    Public Property vccost As List(Of clsVCcost)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vccost = New List(Of clsVCcost)
    End Sub
End Class
