Imports ProjectBoardDefinitions
Public Class clsBewertungWeb
    Public Property key As String
    Public Property bewertung As clsBewertungDB

    Sub New()
        _key = ""
        _bewertung = New clsBewertungDB
    End Sub
End Class
