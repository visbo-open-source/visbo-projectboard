
Public Class clsConstellationItem

    Public Property projectName As String
    Public Property variantName As String
    Public Property Start As Date
    Public Property show As Boolean
    Public Property zeile As Integer

    Sub New()

        _projectName = ""
        _variantName = ""
        _Start = StartofCalendar.AddMonths(-1)
        _show = True
        _zeile = 0

    End Sub
End Class
