Public Class clsWebVC

    Inherits clsWebOutput

    Public Property vc As List(Of clsVC)

    Sub New()
        _vc = New List(Of clsVC)
    End Sub
End Class
