Public Class clsConstellation

    Private _allItems As SortedList(Of String, clsConstellationItem)
    Private _constellationName As String = "Last"

    ''' <summary>
    ''' setzt den Namen; wenn Nothing ode rleer , dann wird als Name Last gesetzt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property constellationName As String
        Get
            constellationName = _constellationName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                If value.Trim.Length > 0 Then
                    _constellationName = value.Trim
                Else
                    _constellationName = "Last"
                End If
            Else
                _constellationName = "Last"
            End If
        End Set
    End Property

    ''' <summary>
    ''' gibt eine komplette Liste an Projekt-Namen zurück, die in der Constellation auftreten;
    ''' by default unabhängig, ob mit Show Attribute oder ohne 
    ''' wenn considerShowAttr = true , dann werden nur die Elemente mit ShowValue gesucht 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectNames(Optional ByVal considerShow As Boolean = False, _
                                             Optional ByVal showValue As Boolean = True) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim pName As String

            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                pName = kvp.Value.projectName
                If Not tmpCollection.Contains(pName) Then
                    tmpCollection.Add(pName, pName)
                End If
            Next

            getProjectNames = tmpCollection

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
            Dim tmpResult As Integer = 0
            Dim atLeastOne As Boolean = False
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems

                If pName = kvp.Value.projectName Then
                    If kvp.Value.variantName = "" Then
                        atLeastOne = True
                    Else
                        tmpResult = tmpResult + 1
                    End If
                End If

            Next

            ' falls es gar nix gibt ... 
            If tmpResult = 0 And Not atLeastOne Then
                tmpResult = -1
            End If

            getVariantZahl = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt die Namen der existierenden Varianten in einer Liste zurück 
    ''' die "leere" Variante wird als () bzw "" zurückgegeben , alle anderen Varianten als (Variante-Name)
    ''' Voraussetzung: _allprojects ist eine sortierte Liste
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getVariantNames(ByVal pName As String, ByVal mitKlammer As Boolean) As Collection
        Get
            Dim tmpCollection As New Collection
            Dim vName As String

            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems

                If pName = kvp.Value.projectName Then
                    If mitKlammer Then
                        vName = "(" & kvp.Value.variantName & ")"
                    Else
                        vName = kvp.Value.variantName
                    End If

                    tmpCollection.Add(vName)

                End If

            Next

            getVariantNames = tmpCollection

        End Get
    End Property

    Public ReadOnly Property Liste() As SortedList(Of String, clsConstellationItem)

        Get
            Liste = _allItems
        End Get

    End Property


    Public ReadOnly Property getItem(key As String) As clsConstellationItem

        Get
            getItem = _allItems(key)
        End Get

    End Property

    Public ReadOnly Property count() As Integer

        Get
            count = _allItems.Count
        End Get

    End Property

    ''' <summary>
    ''' aktualisiert das oder die ShowAttribute gemäß dem Zustand in ShowProjekte
    ''' es wird nur Projekt-Name oder der leere Name (dann alle) übergeben; denn es müssen immer alle Varianten betrachtet werden; 
    ''' ShowProjekte muss vorher aktualsiert worden sein  
    ''' </summary>
    ''' <param name="pName">Projektname, wenn leer - alle behandeln</param>
    ''' <remarks></remarks>
    Public Sub updateShowAttributes(Optional ByVal pName As String = "")
        Dim currentProjectName As String = ""
        Dim hproj As clsProjekt

        ' es werden alle Einträge gemäß Status Showprojekte aktualisiert 
        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
            ' alle bzw. nur den einen Namen behandeln 
            If pName = "" Or pName = kvp.Value.projectName Then

                If ShowProjekte.contains(kvp.Value.projectName) Then
                    hproj = ShowProjekte.getProject(kvp.Value.projectName)
                    kvp.Value.show = (hproj.variantName = kvp.Value.variantName)
                Else
                    kvp.Value.show = False
                End If

            End If
        Next


    End Sub

    Public ReadOnly Property copy(Optional ByVal cName As String = "Copy") As clsConstellation
        Get
            Dim copyResult As New clsConstellation

            With copyResult
                .constellationName = cName

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                    .add(kvp.Value)
                Next

            End With

            copy = copyResult

        End Get
    End Property

    Public Sub add(cItem As clsConstellationItem)

        Dim key As String
        'key = cItem.projectName & "#" & cItem.variantName
        key = calcProjektKey(cItem.projectName, cItem.variantName)
        If Not _allItems.ContainsKey(key) Then
            _allItems.Add(key, cItem)
        End If


    End Sub

    


    ''' <summary>
    ''' löscht den Eintrag mit Schlüssel key; wenn der nicht vorhandenist, dann passiert gar nichts 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub remove(key As String)

        If _allItems.ContainsKey(key) Then
            _allItems.Remove(key)
        End If


    End Sub

    ''' <summary>
    ''' gibt zurück, ob die Constellation die angegebene Variante enthält; 
    ''' wenn withShowFlag = true, dann wird nur True zurückgegeben, wenn die ProjektVariante auch mit Show= true in der Constellation ist
    ''' andernfalls, withShowFlag = false wird nur geprüft, ob die Projekt-Variante in der Konstellation vermerkt ist, unabhängig vom Zustand des Show Attributs  
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <param name="withShowFlag"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function contains(ByVal pvName As String, ByVal withShowFlag As Boolean) As Boolean
        Dim found As Boolean = False
        Dim ix As Integer = 0

        Do While ix <= Me._allItems.Count - 1 And Not found

            If pvName = Me._allItems.ElementAt(ix).Key Then
                If withShowFlag Then
                    If Me._allItems.ElementAt(ix).Value.show = True Then
                        found = True
                    End If
                Else
                    found = True
                End If
            End If

            If Not found Then
                ix = ix + 1
            End If

        Loop

        contains = found
    End Function

    ''' <summary>
    ''' löscht aus dem Szenario alle Einträge von Elementen, die nicht das showAttribute haben 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub reduceToElementsWith(ByVal showAttribute As Boolean)

        Dim toDelete As New Collection

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In Me._allItems
            If kvp.Value.show <> showAttribute Then
                If Not toDelete.Contains(kvp.Key) Then
                    toDelete.Add(kvp.Key, kvp.Key)
                End If

            End If
        Next

        ' jetzt alle Einträge, die nicht das showAttribute trugen, löschen 
        For Each tmpName As String In toDelete

            If Me._allItems.ContainsKey(tmpName) Then
                Me._allItems.Remove(tmpName)
            End If

        Next

    End Sub
    ''' <summary>
    ''' sorgt dafür , dass in der Konstellation alle Projekte mit Name oldNAme mit dem neuen Namen bezeichnet werden 
    ''' </summary>
    ''' <param name="oldPName"></param>
    ''' <param name="newPname"></param>
    ''' <remarks></remarks>
    Public Function rename(ByVal oldPName As String, ByVal newPname As String) As Integer

        Dim toAddItems As New SortedList(Of String, clsConstellationItem)
        Dim toDelete As New Collection

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
            If kvp.Value.projectName = oldPName Then

                Dim tmpConstellationItem As clsConstellationItem = kvp.Value
                Dim key As String = kvp.Key
                ' Vermerk machen zum löschen
                toDelete.Add(key, key)

                ' jetzt das Item neu aufbauen ...
                With tmpConstellationItem
                    .projectName = newPname
                    key = calcProjektKey(.projectName, .variantName)
                End With

                ' Vermerk machen zum Ergänzen 
                toAddItems.Add(key, tmpConstellationItem)

            End If
        Next

        If toDelete.Count <> toAddItems.Count Then
            Call MsgBox("fehler: " & toDelete.Count & ", " & toAddItems.Count)
        End If

        For Each tmpName As String In toDelete
            _allItems.Remove(tmpName)
        Next

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In toAddItems
            _allItems.Add(kvp.Key, kvp.Value)
        Next

        rename = toAddItems.Count

    End Function

    Sub New()

        _allItems = New SortedList(Of String, clsConstellationItem)

    End Sub

    ''' <summary>
    ''' erstellt auf Basis der übergebenen projektliste vom Typ ProjekteAlle eine Konstellation
    ''' wenn kein Name übergeben wird, lautet der Name "Last" 
    ''' wenn keine Angabe zu takeAll gemacht wird, werden sowohl Show als auch noShow ins Szenario aufgenommen 
    ''' </summary>
    ''' <param name="projektListe"></param>
    ''' <remarks></remarks>
    Sub New(ByVal projektListe As clsProjekteAlle, _
            Optional ByVal fullProjectNames As SortedList(Of String, String) = Nothing, _
            Optional ByVal cName As String = "Last", _
            Optional ByVal takeWhat As Integer = ptSzenarioConsider.all)

        _allItems = New SortedList(Of String, clsConstellationItem)
        Me.constellationName = cName

        If IsNothing(projektListe) Then
            ' bereits fertig - es ist eine leere Constelaltion mit Name cNAme
        Else

            If Not IsNothing(fullProjectNames) Then

                Dim newConstellationItem As clsConstellationItem

                For Each kvp As KeyValuePair(Of String, String) In fullProjectNames

                    Dim fullName As String = kvp.Key
                    Dim hproj As clsProjekt = projektListe.getProject(fullName)

                    If Not IsNothing(hproj) Then
                        newConstellationItem = New clsConstellationItem

                        With newConstellationItem
                            .projectName = hproj.name
                            .variantName = hproj.variantName
                            .zeile = 0
                            .Start = hproj.startDate

                            If ShowProjekte.contains(.projectName) Then

                                Dim shownProject As clsProjekt = ShowProjekte.getProject(.projectName)

                                If shownProject.variantName = .variantName Then
                                    .show = True
                                    .zeile = shownProject.tfZeile
                                Else
                                    .show = False
                                End If

                            Else
                                .show = False
                            End If


                        End With

                        ' welche Projekte bzw Projekt-Varianten sollen ins Szenario aufgenommen werden ? 
                        If takeWhat = ptSzenarioConsider.all Then
                            Me.add(newConstellationItem)

                        ElseIf takeWhat = ptSzenarioConsider.show And newConstellationItem.show Then
                            Me.add(newConstellationItem)

                        ElseIf takeWhat = ptSzenarioConsider.noshow And Not newConstellationItem.show Then
                            Me.add(newConstellationItem)
                        End If


                    End If

                Next

            Else

                For Each kvp As KeyValuePair(Of String, clsProjekt) In projektListe.liste

                    Dim newConstellationItem As clsConstellationItem = New clsConstellationItem

                    With newConstellationItem
                        .projectName = kvp.Value.name
                        .variantName = kvp.Value.variantName
                        .zeile = 0
                        .Start = kvp.Value.startDate

                        If ShowProjekte.contains(.projectName) Then

                            Dim shownProject As clsProjekt = ShowProjekte.getProject(.projectName)
                            ' das folgende stellt sicher, dass alle Varianten immer auf der gleichen Zeile sind 
                            .zeile = calcYCoordToZeile(projectboardShapes.getCoord(shownProject.name)(0))
                            If .zeile < 2 Then
                                .zeile = 0
                            End If

                            If shownProject.variantName = .variantName Then
                                .show = True
                            Else
                                .show = False
                            End If

                        Else
                            .show = False
                        End If

                    End With

                    ' welche Projekte bzw Projekt-Varianten sollen ins Szenario aufgenommen werden ? 
                    If takeWhat = ptSzenarioConsider.all Then
                        Me.add(newConstellationItem)

                    ElseIf takeWhat = ptSzenarioConsider.show And newConstellationItem.show Then
                        Me.add(newConstellationItem)

                    ElseIf takeWhat = ptSzenarioConsider.noshow And Not newConstellationItem.show Then
                        Me.add(newConstellationItem)
                    End If

                Next
            End If

        End If

    End Sub

End Class
