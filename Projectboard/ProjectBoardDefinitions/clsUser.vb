Public Class clsUser
    Public Property _id As String
    Public Property email As String
    Public Property name As String
    Public Property _v As Integer
    Public Property posts As String()
    Public Property profile As clsUserProfile
    Public Property created_at As String

    Sub New()
        _id = ""
        _email = "not set"
        _name = "not set"
        _v = 0
        _posts = {"not set", "not yet set!"}
        _profile = New clsUserProfile
        _created_at = ""
    End Sub
End Class
