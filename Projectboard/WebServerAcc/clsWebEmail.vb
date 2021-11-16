Public Class clsWebEmail

    Inherits clsWebOutput
    Public Property mail As clsEmail

    Sub New()
        _mail = New clsEmail
    End Sub
End Class

