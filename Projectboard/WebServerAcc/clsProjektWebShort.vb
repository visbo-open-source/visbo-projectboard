Imports ProjectBoardDefinitions

Public Class clsProjektWebShort
    Public Property _id As String
    Public Property name As String
    Public Property vpid As String
    Public Property timestamp As Date
    Public Property Erloes As Double
    Public Property startDate As Date
    Public Property endDate As Date
    Public Property status As String

    ' ur: 20210915 property vpStatus wird aus dem vp übernommen
    Public Property vpStatus As String
    Public Property variantName As String
    Public Property ampelStatus As String


    Public Sub New()
        _id = ""
        _name = "Project Name"
        _vpid = ""
        _timestamp = Date.MinValue
        _Erloes = 0
        _startDate = Date.MinValue
        _endDate = Date.MaxValue
        'ur: 211202: _status = ProjektStatus(0)
        _vpStatus = ""
        _variantName = ""
        _ampelStatus = ""

    End Sub
End Class
