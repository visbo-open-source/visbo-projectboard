Public Class clsUserReg
    Public Property _id As String
    Public Property email As String
    Public Property profile As clsUserProfile
    Public Property created_at As Date
    Public Property updated_at As Date

    Sub New()
        _id = ""
        _email = "not set"
        _profile = New clsUserProfile
        _created_at = Date.MinValue
        _updated_at = Date.MinValue
    End Sub
End Class
