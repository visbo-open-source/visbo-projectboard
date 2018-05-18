Public Class clsVPLock
    Public Property variantName As String
    Public Property email As String
    Public Property createdAt As Date
    Public Property expiresAt As Date

    Sub New()
        variantName = ""
        email = "someone@visbo.de"
        createdAt = Date.MinValue
        expiresAt = Date.MinValue
    End Sub
End Class
