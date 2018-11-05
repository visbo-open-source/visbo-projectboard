Public Class clsWebVP

    Inherits clsWebOutput
    Public Property vp As List(Of clsVP)

    Sub New()
        _vp = New List(Of clsVP)
    End Sub
End Class
