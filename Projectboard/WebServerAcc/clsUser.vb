Public Class clsUser
    Public Property _id As String
    Public Property email As String
    Public Property role As String


    Sub New()
        _id = ""
        _email = "not known"
        _role = "User/Admin"
    End Sub
End Class
