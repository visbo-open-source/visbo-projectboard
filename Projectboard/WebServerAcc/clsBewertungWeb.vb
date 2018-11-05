Imports ProjectBoardDefinitions
''' <summary>
''' Klassendefinition für eine Bewertung (Phase und/oder Meilenstein) Zugriff über ReST
''' </summary>
Public Class clsBewertungWeb
    Public Property key As String
    Public Property bewertung As clsBewertungDB

    Sub New()
        _key = ""
        _bewertung = New clsBewertungDB
    End Sub
End Class
