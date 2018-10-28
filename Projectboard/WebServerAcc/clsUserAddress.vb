Public Class clsUserAddress
    Public Property street As String
    Public Property city As String
    Public Property zip As String
    Public Property state As String
    Public Property country As String
    Sub New()
        _street = ""
        _city = ""
        _zip = ""
        _state = "Bayern"
        _country = "Deutschland"
    End Sub
End Class
