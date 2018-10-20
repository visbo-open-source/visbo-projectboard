Public Class clsWebVPlock

    Inherits clsWebOutput
    Public Property lock As List(Of clsVPLock)
    Sub New()
        _lock = New List(Of clsVPLock)
    End Sub
End Class
