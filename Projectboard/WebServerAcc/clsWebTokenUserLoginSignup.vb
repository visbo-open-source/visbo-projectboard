Imports System.Globalization
Public Class clsWebTokenUserLoginSignup
    Public Property state As String
    Public Property message As String
    Public Property token As String
    Public Property user As clsUserReg
    Sub New()
        _state = "failure"
        _message = "not successfully logged in"
        _token = ""
        _user = New clsUserReg()
    End Sub
End Class
