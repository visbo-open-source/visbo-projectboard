Public Class clsProjektvorlagen

    Private AllProjects As SortedList(Of String, clsProjektvorlage)


    Public Sub Add(project As clsProjektvorlage)

        Try
            AllProjects.Add(project.VorlagenName, project)
        Catch ex As Exception
            ' wenn das Projekt überschrieben werden muss ...
            AllProjects.Remove(project.VorlagenName)
            AllProjects.Add(project.VorlagenName, project)

        End Try


    End Sub


    Public Sub Remove(projectname As String)

        Try
            AllProjects.Remove(projectname)
        Catch ex As Exception

        End Try



    End Sub

    Public ReadOnly Property Liste() As SortedList(Of String, clsProjektvorlage)
        Get
            Liste = AllProjects
        End Get
    End Property

    Public ReadOnly Property Contains(ByVal name As String) As Boolean
        Get
            Contains = AllProjects.ContainsKey(name)
        End Get
    End Property

    Public ReadOnly Property Count() As Integer

        Get
            Count = AllProjects.Count
        End Get

    End Property

    Public ReadOnly Property getProject(projectname As String) As clsProjektvorlage

        Get
            If AllProjects.ContainsKey(projectname) Then
                getProject = AllProjects.Item(projectname)
            Else
                getProject = Nothing
            End If

        End Get

    End Property

    ''' <summary>
    ''' gibt das Projekt an der Stelle mit Nummer Index zurück
    ''' Index darf 0 bis count-1 sein 
    ''' Wenn Index falsche Werte hat, wird Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(index As Integer) As clsProjektvorlage

        Get
            If index >= 0 Or index <= AllProjects.Count - 1 Then
                getProject = AllProjects.ElementAt(index).Value
            Else
                getProject = Nothing
            End If

        End Get

    End Property


    Public Sub New()

        AllProjects = New SortedList(Of String, clsProjektvorlage)

    End Sub

End Class
