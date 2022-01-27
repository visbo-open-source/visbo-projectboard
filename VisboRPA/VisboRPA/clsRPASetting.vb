Public Class clsRPASetting

    Public VisboCenter As String
    Public VisboUrl As String
    Public VisboConfigFiles As String
    Public activePortfolio As String

    Public Sub New()
        VisboCenter = ""
        VisboUrl = "https://my.visbo.net/api"
        VisboConfigFiles = ""
        activePortfolio = ""
    End Sub

End Class
