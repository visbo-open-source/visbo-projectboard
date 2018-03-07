Public Class clsVC
    Public Property _id As String
    Public Property Name As String
    Public Property Users() As clsVCuser
    Sub New()
        _id = ""
        _Name = "not named"
        _Users = New clsVCUser()
    End Sub
End Class
