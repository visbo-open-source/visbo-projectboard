Public Class clsDependenciesOfPDB

    Public listOfDep1 As SortedList(Of String, clsDependencyDB)
    Public projectName As String
    Public Id As String


    Public Sub copyTo(ByRef projektDependencies As clsDependenciesOfP)

        Dim dependency As clsDependency

        With projektDependencies

            .projectName = Me.projectName

            For Each kvp As KeyValuePair(Of String, clsDependencyDB) In listOfDep1
                dependency = New clsDependency
                kvp.Value.copyTo(dependency)
                projektDependencies.Add(dependency, True)
            Next

        End With

    End Sub

    Public Sub copyFrom(ByVal projektDependencies As clsDependenciesOfP)

        Dim dependency As clsDependencyDB

        projectName = projektDependencies.projectName

        For Each kvp As KeyValuePair(Of String, clsDependency) In projektDependencies.getListe(PTdpndncyType.inhalt)
            dependency = New clsDependencyDB
            dependency.copyFrom(kvp.Value)
            listOfDep1.Add(kvp.Value.dependentProject, dependency)
        Next


    End Sub

    Public Sub New()

        listOfDep1 = New SortedList(Of String, clsDependencyDB)
        projectName = ""

    End Sub

    Public Class clsDependencyDB

        Public project As String
        Public dependentProject As String
        Public type As Integer
        Public degree As Integer
        Public description As String

        Sub copyTo(ByRef dep As clsDependency)

            With dep

                .project = Me.project
                .dependentProject = Me.dependentProject
                .type = Me.type
                .degree = Me.degree
                .description = Me.description

            End With

        End Sub

        Sub copyFrom(ByRef dep As clsDependency)

            With dep

                Me.project = .project
                Me.dependentProject = .dependentProject
                Me.type = .type
                Me.degree = .degree
                Me.description = .description

            End With

        End Sub

        Public Sub New()
            project = ""
            dependentProject = ""
            Type = PTdpndncyType.inhalt
            degree = PTdpndncy.schwach
            description = ""
        End Sub

    End Class

End Class
