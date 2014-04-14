''' <summary>
''' Klasse zur Beschreibung der Abhängigkeiten eines Projektes
''' für jeden Typ von Abhängigkeit wird eine sortierte Liste angelegt
''' aktuell gibt es nur eine Liste: Inhalt
''' </summary>
''' <remarks></remarks>
Public Class clsDependenciesOfP

    Private listOfDep1 As SortedList(Of String, clsDependency)
    Private _projectName As String


    Public Property projectName As String

        Get
            projectName = _projectName
        End Get
        Set(value As String)
            _projectName = value
        End Set

    End Property

    Public ReadOnly Property Count(ByVal type As Integer) As Integer
        Get
            Select Case type

                Case PTdpndncyType.inhalt
                    Count = listOfDep1.Count
                Case Else
                    Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")

            End Select
        End Get
    End Property


    ''' <summary>
    ''' stellt die Liste der Abhängigkeiten zu dem besagten Projekt = Me.projectName bereit 
    ''' </summary>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getListe(ByVal type As Integer) As SortedList(Of String, clsDependency)
        Get

            Select Case type

                Case PTdpndncyType.inhalt
                    getListe = listOfDep1

                Case Else
                    Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")

            End Select

        End Get
    End Property

    ''' <summary>
    ''' gibt Nothing zurück, wenn es für den angegebenen Typ und angegebenes Projekt keine Abhängigkeit gibt 
    ''' Fehler, wenn Typ nicht bekannt ist
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="dependentProjectName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDependency(ByVal type As Integer, ByVal dependentProjectName As String) As clsDependency
        Get

            Select Case type

                Case PTdpndncyType.inhalt
                    If listOfDep1.ContainsKey(dependentProjectName) Then
                        getDependency = listOfDep1.Item(dependentProjectName)
                    Else
                        getDependency = Nothing
                    End If

                Case Else
                    Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")

            End Select

        End Get
    End Property


    ''' <summary>
    ''' fügt der Liste von Projekt-Abhängigkeiten eine neue hinzu
    ''' wenn diese Abhängigkeit bereits existiert: 
    ''' wenn overwrite=true, dann wird die Dependency ggf überschrieben; fungiert damit als "replace"
    ''' </summary>
    ''' <param name="dependency"></param>
    ''' <param name="overwrite">gibt an ob  die Dependency ggf überschreiben werden soll, falls sie bereits existiert</param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal dependency As clsDependency, ByVal overwrite As Boolean)

        Dim key As String

        With dependency
            If .project <> Me.projectName Then
                ' es kann keine Abhängigkeit hinzugefügt werden, die nicht zu Me.projectName gehört
                Throw New ArgumentException("Abhängigkeit gehört nicht zu " & Me.projectName)
            ElseIf .dependentProject = Me.projectName Then
                ' es kann keine Abhängigkeit zu sich selber aufgebaut 
                Throw New ArgumentException("Abhängigkeit zu sich selber nicht zugelassen " & Me.projectName)
            Else
                key = .dependentProject
            End If

        End With


        Select Case dependency.type

            Case PTdpndncyType.inhalt

                Try
                    If Not listOfDep1.ContainsKey(key) Then
                        listOfDep1.Add(key, dependency)
                    ElseIf overwrite Then
                        listOfDep1.Remove(key)
                        listOfDep1.Add(key, dependency)
                    End If
                Catch ex As Exception
                    Throw New ArgumentException("Abhängigkeit kann nicht geschrieben werden ...")
                End Try


            Case Else
                Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")

        End Select

       

    End Sub


    ''' <summary>
    ''' löscht die Abhängigkeit aus der Liste mit Dependencies
    ''' </summary>
    ''' <param name="dependency"></param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal dependency As clsDependency)


        Dim key As String

        With dependency
            If .project <> Me.projectName Then
                ' es kann keine Abhängigkeit gelöscht werden, die nicht zu Me.projectName gehört
                Throw New ArgumentException("Abhängigkeit gehört nicht zu " & Me.projectName)
            Else
                key = .dependentProject
            End If
        End With


        Select dependency.type

            Case PTdpndncyType.inhalt
                Try
                    listOfDep1.Remove(key)
                Catch ex As Exception

                End Try
            Case Else
                Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")

        End Select


       

    End Sub

    ''' <summary>
    ''' löscht die Abhängigkeit mit dem angegebenen Schlüssel aus der Liste von Abhängigkeiten
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal key As String, ByVal type As Integer)


        Select Case type
            Case PTdpndncyType.inhalt

                Try
                    listOfDep1.Remove(key)
                Catch ex As Exception

                End Try

            Case Else
                Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")
        End Select
        

    End Sub

    Public Sub New()

        listOfDep1 = New SortedList(Of String, clsDependency)
        _projectName = ""

    End Sub

    Public Sub New(ByVal projectName As String)

        listOfDep1 = New SortedList(Of String, clsDependency)
        _projectName = projectName

    End Sub
End Class
