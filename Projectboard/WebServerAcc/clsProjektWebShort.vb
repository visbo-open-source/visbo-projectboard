Public Class clsProjektWebShort
    Public Property _id As String
    Public Property name As String
    Public Property vpid As String
    Public Property timestamp As Date
    Public Property Erloes As Double
    Public Property startDate As Date
    Public Property endDate As Date
    Public Property status As String

    Public Property variantName As String
    Public Property ampelStatus As String
    Public Property kundennummer As String

    Public Sub New()
        _id = ""
        _name = "Project Name"
        _vpid = ""
        _timestamp = Date.MinValue
        _Erloes = 0
        _startDate = Date.MinValue
        _endDate = Date.MaxValue
        _status = ""
        _variantName = ""
        _ampelStatus = ""
        _kundennummer = ""
    End Sub
End Class
