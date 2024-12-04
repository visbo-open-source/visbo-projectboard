Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Public Class clsJsonReader
    Public Shared Function ReadJsonFile(Of T)(ByVal filePath As String) As T
        Try
            ' Prüfen ob die Datei existiert
            If Not File.Exists(filePath) Then
                Throw New FileNotFoundException($"Die JSON-Datei wurde nicht gefunden: {filePath}")
            End If

            ' JSON-Datei einlesen und direkt in die Klasse deserialisieren
            Dim jsonContent As String = File.ReadAllText(filePath)
            Dim result As T = JsonConvert.DeserializeObject(Of T)(jsonContent)

            Return result

        Catch ex As JsonSerializationException
            Console.WriteLine($"Fehler bei der Deserialisierung: {ex.Message}")
            Return Nothing
        Catch ex As Exception
            Console.WriteLine($"Ein Fehler ist aufgetreten: {ex.Message}")
            Return Nothing
        End Try
    End Function

End Class
