''' <summary>
''' for importing AI generated tasklists and aggregate and allocate it to a VISBO project 
''' </summary>
Public Class clsProjectPhaseEfforts

    Public projectName As String
    Public alltheEfforts As List(Of clsPhaseEfforts)

    Public Sub New()
        projectName = "AI generated project " & Date.Now.ToShortTimeString
        alltheEfforts = New List(Of clsPhaseEfforts)
    End Sub

End Class
