Imports System.Globalization
Public Class clsWebTokenUserLoginSignup

    Inherits clsWebOutput
    Public Property token As String
    Public Property user As clsUserReg
    Sub New()
        _token = ""
        _user = New clsUserReg()
    End Sub
End Class
