
Imports ProjectBoardDefinitions
Public Class clsAllUnitsDefinitionWeb

    Public name As String
    Public type As Integer
    Public uid As Integer
    Public path As String
    Public entryDate As Date
    Public exitDate As Date
    Public isExternRole As Boolean
    Public defCapaMonth As Double
    ' ? in API so beschrieben  Public defaultCapa As Double
    Public defCapaDay As Double
    Public dailyRate As Double
    Public employeeNr As String
    Public aliases As String()
    Public isAggregationRole As Boolean
    Public isSummaryRole As Boolean

    Public Sub New()

        name = ""
        type = 1
        uid = 0
        path = ""
        entryDate = Date.MinValue.ToUniversalTime
        'exitDate = CDate("31.12.2200").ToUniversalTime
        exitDate = DateAndTime.DateSerial(2200, 12, 31)
        Dim maxDate As Date = Date.MaxValue.ToUniversalTime
        isExternRole = False
        defCapaMonth = -1
        defCapaDay = -1
        dailyRate = 0
        employeeNr = ""
        ' am 10.1. dazugekommen 
        aliases = Nothing
        isAggregationRole = False
        isSummaryRole = False

    End Sub

End Class
