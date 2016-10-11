''' <summary>
''' Klasse für AlleProjekte
''' </summary>
''' <remarks></remarks>
Public Class clsProjekteAlle
    Private _allProjects As SortedList(Of String, clsProjekt)

    Public Sub New()
        _allProjects = New SortedList(Of String, clsProjekt)
    End Sub


    ''' <summary>
    ''' fügt der Sorted List ein Projekt-Element mit Schlüssel key hinzu 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="project"></param>
    ''' <remarks></remarks>
    Public Sub Add(ByVal key As String, ByVal project As clsProjekt)

        _allProjects.Add(key, project)

    End Sub


    ''' <summary>
    ''' gets or sets the sortedlist of (string, clsprojekt)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property liste() As SortedList(Of String, clsProjekt)
        Get
            liste = _allProjects
        End Get

        Set(value As SortedList(Of String, clsProjekt))
            _allProjects = value
        End Set

    End Property

    ''' <summary>
    ''' true, wenn die SortedList ein Element mit angegebenem Key enthält
    ''' false, sonst
    ''' </summary>
    ''' <param name="key"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Containskey(ByVal key As String) As Boolean
        Get
            Containskey = _allProjects.ContainsKey(key)
        End Get
    End Property

   

    ''' <summary>
    ''' gibt die Anzahl Listenelemente der Sorted Liste zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count() As Integer
        Get
            Count = _allProjects.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt das erste Element der Liste zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property First() As clsProjekt
        Get
            If _allProjects.Count > 0 Then
                First = _allProjects.First.Value
            Else
                First = Nothing
            End If
        End Get
    End Property


    ''' <summary>
    ''' gibt eine Liste der vorkommenden Meilenstein Namen in der Menge von Projekte zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getMilestoneNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next

            getMilestoneNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der vorkommenden Phasen-Namen in der Menge der Projekte an ...  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseNames() As Collection

        Get

            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getPhaseNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next


            getPhaseNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Rollen, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getRoleNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next


            getRoleNames = tmpListe
        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Kostenarten, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCostNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpCollection As Collection = kvp.Value.getCostNames

                For Each tmpName As String In tmpCollection
                    If Not tmpListe.Contains(tmpName) Then
                        tmpListe.Add(tmpName, tmpName)
                    End If
                Next

            Next

            getCostNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Business Units, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBUNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpBU As String = kvp.Value.businessUnit
                If Not IsNothing(tmpBU) Then
                    If tmpBU.Trim.Length > 0 Then
                        If Not tmpListe.Contains(tmpBU) Then
                            tmpListe.Add(tmpBU, tmpBU)
                        End If
                    End If
                End If

            Next

            getBUNames = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Projektvorlagen, die in der Menge von Projekten vorkommen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTypNames() As Collection
        Get
            Dim tmpListe As New Collection

            ' neu 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In _allProjects

                Dim tmpTyp As String = kvp.Value.VorlagenName
                If Not IsNothing(tmpTyp) Then
                    If tmpTyp.Trim.Length > 0 Then
                        If Not tmpListe.Contains(tmpTyp) Then
                            tmpListe.Add(tmpTyp, tmpTyp)
                        End If
                    End If
                End If

            Next

            getTypNames = tmpListe

        End Get
    End Property
    ''' <summary>
    ''' gibt die Namen der existierenden Varianten in einer Liste zurück 
    ''' die "leere" Variante wird als () zurückgegeben , alle anderen Varianten als (Variante-Name)
    ''' Voraussetzung: _allprojects ist eine sortierte Liste
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getVariantNames(ByVal pName As String, ByVal mitKlammer As Boolean) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim i As Integer = 0
            Dim found As Boolean = False
            Dim vName As String

            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    found = True
                Else
                    i = i + 1
                End If
            End While

            ' Schreibe alle Varianten in die Ergebnis-Liste tmpCollection
            While i < _allProjects.Count And found

                If _allProjects.ElementAt(i).Value.name = pName Then

                    If mitKlammer Then
                        vName = "(" & _allProjects.ElementAt(i).Value.variantName & ")"
                    Else
                        vName = _allProjects.ElementAt(i).Value.variantName
                    End If

                    tmpCollection.Add(vName)
                    i = i + 1
                Else
                    found = False
                End If

            End While

            getVariantNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt das kleinste Start-Datum zurück, das alle Varianten des Projektes haben 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMinDate(ByVal pName As String) As Date
        Get
            Dim tmpDate As Date = StartofCalendar
            Dim i As Integer = 0
            Dim found As Boolean = False


            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    tmpDate = _allProjects.ElementAt(i).Value.startDate
                    found = True
                    i = i + 1
                Else
                    i = i + 1
                End If
            End While

            ' ist ein Datum einer weiteren Variante kleiner ? 


            While i < _allProjects.Count And found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    If DateDiff(DateInterval.Day, tmpDate, _allProjects.ElementAt(i).Value.startDate) < 0 Then
                        tmpDate = _allProjects.ElementAt(i).Value.startDate
                    End If
                    i = i + 1
                Else
                    found = False
                End If

            End While

            getMinDate = tmpDate

        End Get
    End Property


    ''' <summary>
    ''' gibt das größte Ende-Datum zurück, das alle Varianten des Projekts haben
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMaxDate(ByVal pName As String) As Date
        Get

            Dim tmpDate As Date = StartofCalendar.AddMonths(240)
            Dim i As Integer = 0
            Dim found As Boolean = False


            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    tmpDate = _allProjects.ElementAt(i).Value.endeDate
                    found = True
                    i = i + 1
                Else
                    i = i + 1
                End If
            End While

            ' ist ein Datum einer weiteren Variante größer ? 


            While i < _allProjects.Count And found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    If DateDiff(DateInterval.Day, tmpDate, _allProjects.ElementAt(i).Value.endeDate) > 0 Then
                        tmpDate = _allProjects.ElementAt(i).Value.endeDate
                    End If
                    i = i + 1
                Else
                    found = False
                End If

            End While

            getMaxDate = tmpDate


        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück, das den pName, vName als Projekt- bzw. Varianten-NAme enthält
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(ByVal pName As String, ByVal vName As String) As clsProjekt
        Get
            Dim key As String = calcProjektKey(pName, vName)
            If _allProjects.ContainsKey(key) Then
                getProject = _allProjects(key)
            Else
                getProject = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt das Element zurück, das den angegebenen Schlüssel key enthält
    ''' </summary>
    ''' <param name="key">key = pName#vName</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(ByVal key As String) As clsProjekt
        Get

            If _allProjects.ContainsKey(key) Then
                getProject = _allProjects(key)
            Else
                getProject = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die entsprechende bezeichnete Variante zurück
    ''' VariantNummer = 0 => 1. Projekt-Vorkommen, meist mit Varianten-Namen "" 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="variantNummer"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProject(ByVal pName As String, ByVal variantNummer As Integer) As clsProjekt
        Get


            Dim i As Integer = 0
            Dim found As Boolean = False

            ' Positioniere position auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found

                If _allProjects.ElementAt(i).Value.name = pName Then
                    found = True
                Else
                    i = i + 1
                End If



            End While


            If found Then
                getProject = _allProjects.ElementAt(i + variantNummer).Value
            Else
                getProject = Nothing
            End If


        End Get
    End Property




    ''' <summary>
    ''' gibt die Anzahl Varianten für den übergebenen pName an 
    ''' Das Projekt mit variantName = "" zählt dabei nicht als Variante 
    ''' es gibt nur das Projekt mit Variante "": 0
    ''' es gibt nicht einmal das Projekt mit Namen pName: -1
    ''' Anzahl Varianten mit variantName ungleich "": sonst
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getVariantZahl(ByVal pName As String) As Integer
        Get
            Dim anzahl As Integer = 0
            Dim i As Integer = 0
            Dim found As Boolean = False

            ' Positioniere i auf das erste Vorkommen von pName in der Liste 
            While i < _allProjects.Count And Not found
                If _allProjects.ElementAt(i).Value.name = pName Then
                    found = True
                    anzahl = anzahl + 1
                End If
                i = i + 1

            End While

            ' zähle alle weiteren Vorkommnisse
            While i < _allProjects.Count And found

                If _allProjects.ElementAt(i).Value.name = pName Then
                    anzahl = anzahl + 1
                Else
                    found = False
                End If

                i = i + 1
            End While

            getVariantZahl = anzahl - 1

        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste der unterschiedlichen Projekt-Namen zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectNames() As Collection
        Get
            Dim tmpCollection As New Collection
            Dim i As Integer = 0
            Dim found As Boolean = False
            Dim pName As String

            If Me.Count > 0 Then
                pName = _allProjects.ElementAt(i).Value.name
                tmpCollection.Add(pName)
                i = i + 1

                While i < _allProjects.Count

                    If _allProjects.ElementAt(i).Value.name <> pName Then
                        pName = _allProjects.ElementAt(i).Value.name
                        tmpCollection.Add(pName)
                    End If

                    i = i + 1

                End While

            End If


            getProjectNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' entfernt das Element mit Schlüssel "Key" aus der Sorted List
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub Remove(ByVal key As String)

        If _allProjects.ContainsKey(key) Then
            _allProjects.Remove(key)
        End If

    End Sub

    ''' <summary>
    ''' entfernt alle Projekt-Varianten mit ProjektNamen = pName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <remarks></remarks>
    Public Sub RemoveAllVariantsOf(ByVal pName As String)

        Dim i As Integer = 0
        Dim found As Boolean = False

        ' Positioniere i auf das erste Vorkommen von pName in der Liste 
        While i < _allProjects.Count And Not found
            If _allProjects.ElementAt(i).Value.name = pName Then
                found = True
            Else
                i = i + 1
            End If
        End While

        ' Lösche alle Varianten mit ProjektName = pName 
        While found

            If i < _allProjects.Count Then

                If _allProjects.ElementAt(i).Value.name = pName Then
                    _allProjects.RemoveAt(i)
                Else
                    found = False
                End If

            Else
                found = False
            End If

        End While

    End Sub

    ''' <summary>
    ''' setzt die Liste der Projekte zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()

        _allProjects.Clear()

    End Sub

End Class
