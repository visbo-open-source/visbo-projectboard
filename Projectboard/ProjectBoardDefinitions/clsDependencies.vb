''' <summary>
''' enthält alle Abhängigkeiten der aktuell geladenen Projekte
''' </summary>
''' <remarks></remarks>
Public Class clsDependencies

    Private listOfProjDependencies As SortedList(Of String, clsDependenciesOfP)

    ''' <summary>
    ''' fügt die Liste der Projekt-Abhängigkeiten eines Projektes hinzu, bei Overwrite = true wird überschrieben , falls es diese Liste bereits gibt
    ''' </summary>
    ''' <param name="liste"></param>
    ''' <param name="overwrite"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal liste As clsDependenciesOfP, ByVal overwrite As Boolean)

        Dim key As String

        If Not IsNothing(liste) Then
            key = liste.projectName
            If key <> "" Then
                If listOfProjDependencies.ContainsKey(key) Then
                    If overwrite Then
                        listOfProjDependencies.Remove(key)
                        listOfProjDependencies.Add(key, liste)
                    Else
                        Throw New ArgumentException("Liste existiert schon " & key)
                    End If
                Else
                    listOfProjDependencies.Add(key, liste)
                End If
            Else
                Throw New ArgumentException("key ist leer ! ")
            End If
            
        End If
        

    End Sub

    ''' <summary>
    ''' erzeugt aus den übergebenen Werten eine neue Abhängigkeit ; sofern die schon existiert wird sie in den Ausprägungen 
    ''' degree und description überschrieben  
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="dpName"></param>
    ''' <param name="type"></param>
    ''' <param name="degree"></param>
    ''' <param name="description"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal pName As String, ByVal dpName As String, ByVal type As Integer, _
                       ByVal degree As Integer, ByVal description As String)

        Dim tmpDependency As clsDependency
        Dim tmpDependenciesOfP As clsDependenciesOfP

        tmpDependency = Me.getDependency(type, pName, dpName)

        If Not IsNothing(tmpDependency) Then
            With tmpDependency
                .description = description
                If degree = PTdpndncy.schwach Or degree = PTdpndncy.stark Then
                    .degree = degree
                End If
            End With
        Else
            ' Erzeugen der abhängigkeit 
            tmpDependency = New clsDependency(pName, dpName, type, degree, description)

            ' Gibt es bereits für pName eine Liste von abhängigkeiten ?
            If Me.listOfProjDependencies.ContainsKey(pName) Then
                tmpDependenciesOfP = Me.listOfProjDependencies.Item(pName)
                tmpDependenciesOfP.Add(tmpDependency, True)
                ' jetzt fertig ...
            Else
                tmpDependenciesOfP = New clsDependenciesOfP(pName)
                tmpDependenciesOfP.Add(tmpDependency, True)
                ' jetzt noch nicht fertig, da die Liste noch hinzugefügt werden muss 
                Me.Add(tmpDependenciesOfP, True)
            End If



        End If


    End Sub

    ''' <summary>
    ''' gibt Zugriff auf die SortedList of Projekt-Abhängigkeiten 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSortedList() As SortedList(Of String, clsDependenciesOfP)
        Get
            getSortedList = listOfProjDependencies
        End Get
    End Property


    ''' <summary>
    ''' gibt für das angegebene Projekt und den angegebenen Typ die Liste der abhängigen Projekte zurück 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDependenciesOfP(ByVal pName As String) As clsDependenciesOfP
        Get


            If listOfProjDependencies.ContainsKey(pName) Then
                getDependenciesOfP = listOfProjDependencies.Item(pName)
            Else
                getDependenciesOfP = Nothing
            End If


        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl von Projekten zurück, für die es eine Liste von Abhängigkeiten gibt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property projectCount As Integer
        Get
            projectCount = listOfProjDependencies.Count
        End Get
    End Property

    Public ReadOnly Property totalCount As Integer
        Get
            Dim tmpValue As Integer = 0

            For p = 1 To listOfProjDependencies.Count
                ' wenn es mehrere Typen gibt, muss hier eine Schleife durch die Enumeration gemacht werden
                tmpValue = tmpValue + listOfProjDependencies.ElementAt(p - 1).Value.Count(PTdpndncyType.inhalt)
            Next

            totalCount = tmpValue
        End Get
    End Property


    ''' <summary>
    ''' löscht das angegebene Projekt aus allen Listen, wo es ein "dependent Project" ist 
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <remarks></remarks>
    Public Sub removeInComing(ByVal projectName As String)


        For Each kvp As KeyValuePair(Of String, clsDependenciesOfP) In listOfProjDependencies

            If kvp.Value.getListe(PTdpndncyType.inhalt).ContainsKey(projectName) Then
                kvp.Value.Remove(projectName, PTdpndncyType.inhalt)
            End If
            ' hier müssten jetzt die anderen Types ergänzt werden ...

        Next
    End Sub

    ''' <summary>
    ''' liefert die Abhängigkeits-Beziehung zwischen projectName und dependenProjectName zurück, wenn sie existiert
    ''' Nothing, wenn sie nicht existiert
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="projectName"></param>
    ''' <param name="dependentProjectName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDependency(ByVal type As Integer, ByVal projectName As String, ByVal dependentProjectName As String) As clsDependency
        Get
            Dim projectDep As clsDependenciesOfP
            Select Case type

                Case PTdpndncyType.inhalt

                    If listOfProjDependencies.ContainsKey(projectName) Then

                        projectDep = listOfProjDependencies.Item(projectName)
                        getDependency = projectDep.getDependency(type, dependentProjectName)

                    Else
                        getDependency = Nothing
                    End If

                Case Else
                    Throw New ArgumentException("nicht unterstützter Abhängigkeits-Typ")

            End Select

        End Get
    End Property


    ''' <summary>
    ''' bestimmt den Index, der sich aus der Summe der Wertigkeiten der abhängigen Projekte ergibt 
    ''' </summary>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property activeIndex(ByVal projectName As String, ByVal type As Integer) As Integer
        Get
            Dim tmpSum As Integer = 0
            Dim hdependencies As SortedList(Of String, clsDependency)
            ' jetzt werden die Degrees der Abhängigkeit von allen Projekten aufsummiert, die von 
            ' projectName abhängig sind 

            If listOfProjDependencies.ContainsKey(projectName) Then
                hdependencies = listOfProjDependencies.Item(projectName).getListe(type)
                For i = 1 To hdependencies.Count
                    tmpSum = tmpSum + hdependencies.ElementAt(i - 1).Value.degree
                Next
            End If

            activeIndex = tmpSum

        End Get
    End Property

    Public ReadOnly Property activeNumber(ByVal projectName As String, ByVal type As Integer) As Integer
        Get

            Dim tmpSum As Integer = 0

            If listOfProjDependencies.ContainsKey(projectName) Then

                tmpSum = listOfProjDependencies.Item(projectName).getListe(type).Count
                
            End If

            activeNumber = tmpSum

        End Get
    End Property

    ''' <summary>
    ''' es werden die Projekt-Namen ermittelt, die von projectname abhängig sind 
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property activeListe(ByVal projectName As String, ByVal type As Integer) As Collection
        Get
            Dim tmpColl As New Collection
            Dim key As String
            Dim hdependencies As SortedList(Of String, clsDependency)
            ' jetzt werden die Namen aller Projekten in tmpColl geschrieben, die von 
            ' projectName abhängig sind 

            If listOfProjDependencies.ContainsKey(projectName) Then
                hdependencies = listOfProjDependencies.Item(projectName).getListe(type)

                For i = 1 To hdependencies.Count
                    key = hdependencies.ElementAt(i - 1).Key
                    Try
                        tmpColl.Add(key, key)
                    Catch ex As Exception

                    End Try

                Next
            End If

            activeListe = tmpColl
        End Get
    End Property


    ''' <summary>
    ''' bestimmt den passiven Index: das heißt wie oft ist dieses Projekt das dependent Project und mit welchem Grad ...
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property passiveIndex(ByVal projectName As String, ByVal type As Integer) As Integer
        Get
            Dim tmpSum As Integer = 0

            For Each kvp As KeyValuePair(Of String, clsDependenciesOfP) In listOfProjDependencies

                If kvp.Value.getListe(type).ContainsKey(projectName) Then
                    tmpSum = tmpSum + kvp.Value.getListe(type).Item(projectName).degree
                End If

            Next

            passiveIndex = tmpSum
        End Get
    End Property

    ''' <summary>
    ''' bestimmt die Anzahl Projekte, von denen projectName abhängig ist 
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property passiveNumber(ByVal projectName As String, ByVal type As Integer) As Integer

        Get
            Dim tmpSum As Integer = 0

            For Each kvp As KeyValuePair(Of String, clsDependenciesOfP) In listOfProjDependencies

                If kvp.Value.getListe(type).ContainsKey(projectName) Then
                    tmpSum = tmpSum + 1
                End If

            Next

            passiveNumber = tmpSum
        End Get

    End Property

    ''' <summary>
    ''' ' es wird die Liste erstellt, von welchen Projekten projectname abhängig ist
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property passiveListe(ByVal projectName As String, ByVal type As Integer) As Collection
        Get
            Dim tmpColl As New Collection
            Dim key As String

            For Each kvp As KeyValuePair(Of String, clsDependenciesOfP) In listOfProjDependencies
                ' es wird die Liste erstellt, von welchen Projekten projectname abhängig ist 
                If kvp.Value.getListe(type).ContainsKey(projectName) Then
                    key = kvp.Key
                    Try
                        tmpColl.Add(key, key)
                    Catch ex As Exception

                    End Try

                End If

            Next

            passiveListe = tmpColl

        End Get
    End Property

    ''' <summary>
    ''' löscht die Liste der Abhängigkeiten für das angegebene Projekt
    ''' </summary>
    ''' <param name="projectName"></param>
    ''' <remarks></remarks>
    Public Sub remove(ByVal projectName As String)

        Try
            listOfProjDependencies.Remove(projectName)
        Catch ex As Exception

        End Try

    End Sub

    Sub New()

        listOfProjDependencies = New SortedList(Of String, clsDependenciesOfP)

    End Sub
End Class
