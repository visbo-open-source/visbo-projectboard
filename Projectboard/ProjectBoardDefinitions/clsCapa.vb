''' <summary>
''' represents the capa per month of one role in one year
''' </summary>
Public Class clsCapa
    Public Property _id As String
    Public Property vcid As String
    Public Property roleID As String
    Public Property startOfYear As Date
    Public Property capaPerMonth As List(Of Double)


    Public Sub New()
        _id = ""
        _vcid = ""
        _roleID = ""
        _startOfYear = Date.Now
        _capaPerMonth = New List(Of Double)
    End Sub
End Class
