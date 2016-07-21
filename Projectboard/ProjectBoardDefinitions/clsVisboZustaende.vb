''' <summary>
''' enthält bestimme Zustands-Variablen der Projekt-Tafel 
''' </summary>
''' <remarks></remarks>
Public Class clsVisboZustaende
    Public Property showTimeZoneBalken As Boolean
    Public Property projectBoardMode As Integer

    Sub New()
        _showTimeZoneBalken = False
        _projectBoardMode = ptModus.graficboard
    End Sub
End Class
