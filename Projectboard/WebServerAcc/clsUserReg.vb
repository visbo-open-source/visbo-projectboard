Public Class clsUserReg
    Public Property _id As String
    Public Property email As String
    Public Property name As String
    Public Property profile As clsUserProfile
    Public Property created_at As Date
    Public Property updated_at As Date

    Sub New()
        _id = ""
        _email = "not set"
        _name = "not set"
        _profile = New clsUserProfile
        _created_at = Convert.ToDateTime("2018-03-02T16:36:49.122Z")
        _updated_at = Convert.ToDateTime("2018-03-16T16:36:49.122Z")
    End Sub
End Class
