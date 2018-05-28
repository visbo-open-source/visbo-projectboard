Public Class clsVPLock
    Public Property variantName As String
    Public Property email As String
    Public Property createdAt As Date
    Public Property expiresAt As Date

    Sub New()
        _variantName = ""
        _email = "someone@visbo.de"
        _createdAt = Date.MinValue
        _expiresAt = Date.MinValue
    End Sub
End Class
