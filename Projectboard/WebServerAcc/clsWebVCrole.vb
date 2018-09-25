Public Class clsWebVCrole

    Public Property state As String
    Public Property message As String
    Public Property vcrole As List(Of clsVCrole)

    Sub New()
        _state = "unknown"
        _message = "not yet any response"
        _vcrole = New List(Of clsVCrole)
    End Sub
End Class
