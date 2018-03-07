Imports System.Globalization
Public Class clsTokenUserSignIn
    Public Property state As String
    Public Property message As String
    Public Property user As clsUser
    Sub New()
        _state = "failure"
        _message = "not successfully logged in"
        _user = New clsUser()
    End Sub
End Class
