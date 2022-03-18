Public Class clsWebVCUser

    Inherits clsWebOutput

    Public Property user As List(Of clsUserReg)
    Public Property count As Integer


    Sub New()
        _user = New List(Of clsUserReg)
        _count = 0
    End Sub
End Class
