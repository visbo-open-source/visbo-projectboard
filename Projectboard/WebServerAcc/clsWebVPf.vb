Public Class clsWebVPf

    Inherits clsWebOutput
    Public Property vpf As List(Of clsVPf)

    Sub New()
        _vpf = New List(Of clsVPf)
    End Sub
End Class
