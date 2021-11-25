Public Class clsRPASetting

    Public VisboCenter As String
    Public VisboUrl As String
    Public activePortfolio As String

    Public Sub New()
        VisboCenter = ""
        VisboUrl = "https://my.visbo.net/api"
        activePortfolio = ""
    End Sub

End Class
