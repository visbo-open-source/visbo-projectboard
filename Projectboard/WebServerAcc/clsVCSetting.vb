Public Class clsVCSetting
    Public Property _id As String
    Public Property vcid As String
    Public Property name As String
    Public Property userId As String
    Public Property type As String
    Public Property timestamp As Date


    Sub New()
        _id = ""
        _vcid = "not yet defined"
        _name = "no setting name"
        _userId = "not defined"
        _type = "type of setting"
        _timestamp = Date.MinValue

    End Sub


End Class
