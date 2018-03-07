Public Class clsUser
    Public Property _id As String
    Public Property email As String
    Public Property name As String
    Public Property _v As Integer
    Public Property profile As clsUserProfile
    Public Property created_at As String
    Public Property updated_at As String

    Sub New()
        _id = ""
        _email = "not set"
        _name = "not set"
        _v = 0
        _profile = New clsUserProfile
        _created_at = ""
        _updated_at = ""
    End Sub
End Class
