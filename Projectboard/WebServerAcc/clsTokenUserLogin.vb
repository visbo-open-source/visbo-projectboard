Imports Microsoft.Office.Interop.Excel
Imports System.Globalization
Public Class clsTokenUserLogin
    Public Property state As String
    Public Property message As String
    Public Property token As String
    Public Property user As clsUser
    Sub New()
        _state = "failure"
        _message = "not successfully logged in"
        _token = ""
        user = New clsUser()
    End Sub
End Class
