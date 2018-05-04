Public Class clsUserProfile
    Public Property firstname As String
    Public Property lastname As String
    Public Property phone As String
    Public Property company As String
    Public Property address As clsUserAddress
    Sub New()
        _firstname = "NN"
        _lastname = "NN"
        _phone = "not set"
        _company = "not set"
        _address = New clsUserAddress
    End Sub
End Class
