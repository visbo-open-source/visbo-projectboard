Public Class clsWebVCrole

    Inherits clsWebOutput
    Public Property vcrole As List(Of clsVCrole)

    Sub New()
        _vcrole = New List(Of clsVCrole)
    End Sub
End Class
