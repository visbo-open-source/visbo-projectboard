Public Class clsWebVCcost

    Inherits clsWebOutput
    Public Property vccost As List(Of clsVCcost)

    Sub New()
        _vccost = New List(Of clsVCcost)
    End Sub
End Class
