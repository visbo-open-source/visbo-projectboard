Public Class clsJIRA_Task

    Public Property projectName As String
    Public Property Jira_ID As String
    Public Property Vorgangstyp As String
    Public Property Zusammenfassung As String
    Public Property zugewPerson As String
    Public Property Autor As String
    Public Property Prio As String
    Public Property TaskStatus As String
    Public Property loesung As String
    Public Property Erstellt As Date
    Public Property aktualisiert As Date
    Public Property fällig As Date
    Public Property StartDate As Date
    Public Property verknüpfte_JiraID As String
    Public Property Area As String
    Public Property parent_JiraID As String
    Public Property Fortschritt As Integer
    Public Property StoryPoints As Double
    Public Property erledigt As Date
    Public Property SprintName As String
    Public Property SprintStartDate As Date
    Public Property SprintEndDate As Date
    Public Property SprintCompleteDate As Date
    Public Property SprintGoal As String

    Public Sub New()

        _projectName = ""
        _Jira_ID = ""
        _Vorgangstyp = ""
        _Zusammenfassung = ""
        _zugewPerson = ""
        _Autor = ""
        _Prio = ""
        _TaskStatus = ""
        _loesung = ""
        _Erstellt = Date.MinValue
        _aktualisiert = Date.MinValue
        _fällig = Date.MinValue
        _StartDate = Date.MinValue
        _verknüpfte_JiraID = ""
        _parent_JiraID = ""
        _Fortschritt = 0
        _StoryPoints = 0.0
        _erledigt = Date.MinValue
        _SprintName = ""
        _SprintStartDate = Date.MinValue
        _SprintEndDate = Date.MinValue
        _SprintCompleteDate = Date.MinValue
    End Sub

End Class


Public Class clsJIRA_sprint
    Public Property SprintName As String
    Public Property SprintStartDate As Date
    Public Property SprintEndDate As Date
    Public Property SprintCompleteDate As Date
    Public Property SprintGoal As String
    Public Property SprintTasks As SortedList(Of String, String)
    Public Sub New()
        _SprintName = ""
        _SprintStartDate = Date.MinValue
        _SprintEndDate = Date.MaxValue
        _SprintCompleteDate = Date.MaxValue
        _SprintGoal = ""
        _SprintTasks = New SortedList(Of String, String)
    End Sub
End Class
