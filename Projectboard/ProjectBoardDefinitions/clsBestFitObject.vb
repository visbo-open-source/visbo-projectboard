Public Class clsBestFitObject

    Public Property type As String
    Public Property name As String
    Public Property myCollection As Collection
    Public Property isOn As Boolean

    Sub New()
        myCollection = New Collection
        _isOn = False
        _name = ""
        _type = ""
    End Sub
End Class
