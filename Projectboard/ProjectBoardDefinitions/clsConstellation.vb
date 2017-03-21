Public Class clsConstellation

    ''' <summary>
    ''' _allItems ist sorted list with key containing projectName#variantName
    ''' </summary>
    ''' <remarks></remarks>
    Private _allItems As SortedList(Of String, clsConstellationItem)

    ' sortierte Liste eines beliebig zu erstellenden Keys und dem pvName  
    ''' <summary>
    ''' _sortlist is sorted list providing the sequence, in which projects shall be shown in window / multiproject board
    ''' </summary>
    ''' <remarks></remarks>
    Private _sortList As SortedList(Of String, String)

    ''' <summary>
    ''' _lastCustomList merkt sich die letzten Custom Werte, so dass hierhin problemlos zurückgegangen werden kann
    ''' ohne nochmal alles händisch machen zu müssen 
    ''' </summary>
    ''' <remarks></remarks>
    Private _lastCustomList As SortedList(Of String, String)

    ' gibt an, nach welchem Sortierkriterium die _sortList aufgebaut wurde 
    ' 0: alphabetisch nach Name
    ' 1: custom tfzeile 
    ' 2: custom Liste
    ' 3: BU, ProjektStart, Name
    ' 4: Formel: strategic Fit* 100 - risk*90 + 100*Marge + korrFaktor
    Private _sortType As Integer

    Private _constellationName As String = "Last"

    ''' <summary>
    ''' gibt die Zeile zurück, auf der dieses Projekt gezeichnet werden soll 
    ''' 
    ''' </summary>
    ''' <param name="pName">NAme des Projektes, ohne den Varianten-Namen Anteil </param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBoardZeile(ByVal pName As String) As Integer
        Get
            Dim found As Boolean = False
            Dim ix As Integer = 0
            Dim bzeile As Integer = 0

            Do While ix <= _sortList.Count - 1 And Not found
                Dim vglName As String = _sortList.ElementAt(ix).Value
                If vglName = pName Then
                    found = True
                Else
                    ix = ix + 1
                    If ShowProjekte.contains(vglName) Then
                        bzeile = bzeile + 1
                    End If
                End If
            Loop

            getBoardZeile = bzeile + 2

        End Get
    End Property

    ''' <summary>
    ''' liest bzw. schreibt die Sortier-Liste, die die Reihenfolge Information bereitstellt, in der die Projekte dargestellt werden sollen 
    ''' beim Setzen muss ein Parameter mitgegeben, der angibt, um welche Sortierungs-Liste es sich hierbei handelt  
    ''' </summary>
    ''' <param name="sType"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property sortListe(Optional ByVal sType As Integer = ptSortCriteria.alphabet) As SortedList(Of String, String)
        Get
            sortListe = _sortList
        End Get
        Set(value As SortedList(Of String, String))
            Dim correct As Boolean = False
            If Not IsNothing(value) Then
                If value.Count = Me.getProjectNames.Count Then
                    correct = True
                    For Each kvp As KeyValuePair(Of String, String) In value
                        ' prüfen, ob die Gesamt-Liste irgendeine Projekt-Variante mit diesem Projekt-Namen enthält 
                        If Me.containsProject(kvp.Key) Then
                            ' alles in Ordnung 
                        Else
                            ' nicht in Ordnung 
                            correct = False
                        End If
                    Next
                End If
            End If
            If correct Then
                _sortType = sType
                _sortList = New SortedList(Of String, String)

                ' kopieren der Liste 
                For Each kvp As KeyValuePair(Of String, String) In value
                    ' Prüfung auf Enthaltensein kann hier entfallen, da eine sortierte Liste wie value nur eindeutige keys enthalten kann  
                    _sortList.Add(kvp.Key, kvp.Value)
                Next

            End If
        End Set
    End Property

    ''' <summary>
    ''' liest bzw. setzt die lastCustomlist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property lastCustomList As SortedList(Of String, String)
        Get
            lastCustomList = _lastCustomList
        End Get
        Set(value As SortedList(Of String, String))
            Dim correct As Boolean = False
            If Not IsNothing(value) Then
                If value.Count = Me.getProjectNames.Count Then
                    correct = True
                    For Each kvp As KeyValuePair(Of String, String) In value
                        ' prüfen, ob die Gesamt-Liste irgendeine Projekt-Variante mit diesem Projekt-Namen enthält 
                        If Me.containsProject(kvp.Key) Then
                            ' alles in Ordnung 
                        Else
                            ' nicht in Ordnung 
                            correct = False
                        End If
                    Next
                End If
            End If
            If correct Then

                _lastCustomList = New SortedList(Of String, String)

                ' kopieren der Liste 
                For Each kvp As KeyValuePair(Of String, String) In value
                    ' Prüfung auf Enthaltensein kann hier entfallen, da eine sortierte Liste wie value nur eindeutige keys enthalten kann  
                    _lastCustomList.Add(kvp.Key, kvp.Value)
                Next

            End If
        End Set
    End Property
   
    ''' <summary>
    ''' baut eine sortierte Liste der Projekt-Namen auf !
    ''' die Position auf der Projekt-Tafel bzw. im Portfolio Browser ergibt sich dann 
    ''' aus dem Index in der sortierten Liste  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property sortCriteria As Integer

        Get
            sortCriteria = _sortType
        End Get
        Set(value As Integer)

            Dim key As String = ""

            If Not IsNothing(value) Then

                Dim istNull As Boolean = IsNothing(Me._sortList)
                Dim istLeer As Boolean
                If Not istNull Then
                    istLeer = (Me._sortList.Count = 0)
                Else
                    istLeer = True
                End If

                If _sortType <> value Or istNull Or istLeer Then
                    ' nur wenn es unterschiedlich oder wenn es noch gar nicht gesetzt ist, muss etwas getan werden 

                    ' wenn der aktuelle _sortType = CustomTF ist, dann merken der Liste 
                    If _sortType = ptSortCriteria.customTF Then
                        _lastCustomList = _sortList
                    End If

                    ' jetzt müssen die Sort-Keys gesetzt werden 
                    _sortType = value
                    _sortList = New SortedList(Of String, String)
                    For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems

                        Dim hproj As clsProjekt = AlleProjekte.getProject(kvp.Key)
                        If Not IsNothing(hproj) Then

                            ' nur wenn es nicht bereits in sortList enthalten ist 
                            If Not _sortList.ContainsValue(hproj.name) Then

                                hproj = getSortRelevantProject(hproj.name)

                                If Not IsNothing(hproj) Then
                                    key = hproj.getSortKeyForConstellation(_sortType)

                                    If Not _sortList.ContainsKey(key) Then
                                        _sortList.Add(key, hproj.name)
                                    Else
                                        ' es muss ein . ergänzt werden 
                                        key = key & "."
                                        Do While _sortList.ContainsKey(key)
                                            key = key & "."
                                        Loop
                                        _sortList.Add(key, hproj.name)
                                    End If
                                End If
                            End If


                        End If

                            
                    Next

                End If

            End If


        End Set
    End Property

    ''' <summary>
    ''' liefert zu einem gegebenen Projekt-Namen das Projekt ab, das für die Sortier-Schlüssel-Berechnung verwendet werden soll 
    ''' das relevante Projekt ist das, was im Show ist bzw das was als erstes in der Variant-Liste steht  
    ''' Nothing, wenn es das Projekt gar nicht gibt 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getSortRelevantProject(ByVal pName As String) As clsProjekt

        Dim hproj As clsProjekt = Nothing

        If ShowProjekte.contains(pName) Then
            hproj = ShowProjekte.getProject(pName)
        Else
            ' bestimme das hproj, das als erste Variante vorkommt 
            Dim vName As String = ""
            Dim tmpCollection As Collection = Me.getVariantNames(pName, False)
            If Not IsNothing(tmpCollection) Then
                vName = CStr(tmpCollection.Item(1))
            End If
            Dim tmpKey As String = calcProjektKey(pName, vName)
            hproj = AlleProjekte.getProject(tmpKey)
        End If

        getSortRelevantProject = hproj

    End Function

    ''' <summary>
    ''' setzt den Namen; wenn Nothing oder leer , dann wird als Name Last gesetzt 
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
                    If value = "Last" Then
                        _constellationName = "Last" & dbUsername
                    Else
                        _constellationName = value.Trim
                    End If

                Else
                    _constellationName = "Last" & dbUsername
                End If
            Else
                _constellationName = "Last" & dbUsername
            End If
        End Set
    End Property

    Public Sub checkAndCorrectYourself()

        ' Check 1: 
        ' sind alle ShowProjekte auch in der Constellation aufgeführt ? 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            Dim key As String = calcProjektKey(kvp.Value)
            If _allItems.ContainsKey(key) Then
                If _allItems.Item(key).show = True Then
                    ' alles in Ordnung 
                Else
                    Call MsgBox("hat kein Show-Attribut:" & key)
                End If

            Else
                Call MsgBox("Show-Projekt nicht enthalten: " & key)
            End If

        Next

        ' Check 2: 
        ' sind alle Items aus der Constellation mit Attribut Show=true auch in ShowProjekte? 
        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
            If kvp.Value.show = True Then
                Dim hproj As clsProjekt = ShowProjekte.getProject(kvp.Value.projectName)
                If Not IsNothing(hproj) Then
                    If hproj.variantName = kvp.Value.variantName Then
                        ' alles in Ordnung 
                    Else
                        Call MsgBox("hproj ist mit falschem Variant-Name in der Constellation ... " & kvp.Key)
                    End If
                Else
                    Call MsgBox("Item ist nicht in ShowProjekte ... " & kvp.Key)
                End If
            End If

        Next

    End Sub
    ''' <summary>
    ''' setzt in Abhängigkeit von type die Tfzeilen in den clsConstellationItems  
    ''' 
    ''' </summary>
    ''' <param name="sortierTypus"></param>
    ''' <remarks></remarks>
    Public Sub setTfZeilen(ByVal sortierTypus As Integer)

        Dim zeile As Integer = 2
        'Dim sortierListe As SortedList(Of Double, String)

        Select Case sortierTypus
            Case 0
                ' sortiert nach dem Key, also pName#VariantName 
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                    If kvp.Value.show Then
                        kvp.Value.zeile = zeile
                        zeile = zeile + 1
                    Else
                        kvp.Value.zeile = 0
                    End If
                Next
            Case 1
            Case 2
            Case Else

        End Select

    End Sub


    ''' <summary>
    ''' provides a complete list of project names in the current constellation 
    ''' by default: independent of having show-Attribute or not
    ''' when considerShowAttr = true , only names with show-attribute = showvalue are in the output list 
    ''' which sortcriteria shall be applied; default = alphabetical order 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectNames As Collection
        Get
            Dim tmpCollection As New Collection
            Dim pName As String


            ' jetzt is _sortList auf alle Fälle in der richtigen Form ... 
            For Each kvp As KeyValuePair(Of String, String) In _sortList
                pName = kvp.Value

                If Not tmpCollection.Contains(pName) Then
                    tmpCollection.Add(Item:=pName, Key:=pName)
                End If

            Next

            getProjectNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl Varianten für den übergebenen pName an 
    ''' Das Projekt mit variantName = "" zählt dabei auch als Variante 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getVariantZahl(ByVal pName As String) As Integer
        Get
            Dim tmpResult As Integer = 0
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems

                If pName = kvp.Value.projectName Then
                    tmpResult = tmpResult + 1
                End If

            Next

            getVariantZahl = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' provides the list of variant Names in alphabetical order 
    ''' if mitKlammer = true then items will enclosed by ()
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
    ''' ShowProjekte muss vorher aktualisiert worden sein  
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
                    ' jede Variante soll ja in der gleichen Zeile gezeichnet werden ...
                    kvp.Value.zeile = hproj.tfZeile

                    If (hproj.variantName = kvp.Value.variantName) Then
                        kvp.Value.show = True
                    Else
                        kvp.Value.show = False
                    End If

                Else
                    kvp.Value.show = False
                    kvp.Value.zeile = 0
                End If

            End If
        Next


    End Sub


    ''' <summary>
    ''' kopiert eine Constellation, d.h jetzt müssen auch sortType und sortList kopiert werden 
    ''' </summary>
    ''' <param name="cName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property copy(Optional ByVal cName As String = "Last") As clsConstellation
        Get
            Dim copyResult As New clsConstellation

            ' wenn Last, soll es auf den User selber angewendet werden 
            If cName = "Last" Then
                cName = cName & dbUsername
            End If

            With copyResult
                .constellationName = cName

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                    Dim copiedItem As clsConstellationItem = kvp.Value.copy
                    .add(copiedItem)
                Next

                ' jetzt sortliste und sorttype kopieren 
                .sortListe(_sortType) = _sortList

                ' jetzt ggf die lastCustomList kopieren 
                If Not IsNothing(_lastCustomList) Then
                    If _lastCustomList.Count > 0 Then
                        .lastCustomList = _lastCustomList
                    End If
                End If
                
            End With

            copy = copyResult

        End Get
    End Property

    ''' <summary>
    ''' fügt ein clsConstellationItem hinzu und aktualisiert auch die Sortlist entsprechend ... 
    ''' Voraussetzung: in AlleProjekte ist das im Item beschriebene Objekt bereits enthalten 
    ''' im add muss kein Update der lastCustomlist erfolgen, nur beim Remove ... 
    ''' </summary>
    ''' <param name="cItem"></param>
    ''' <remarks></remarks>
    Public Sub add(cItem As clsConstellationItem)

        Dim key As String
        Dim sortKey As String
        'key = cItem.projectName & "#" & cItem.variantName
        key = calcProjektKey(cItem.projectName, cItem.variantName)

        If Not _allItems.ContainsKey(key) Then

            _allItems.Add(key, cItem)

            ' jetzt auch in sortlist aktualisieren 
            If _sortList.ContainsValue(cItem.projectName) And cItem.show Then
                ' Remove den bisherigen Schlüssel 
                Dim ix As Integer = _sortList.IndexOfValue(cItem.projectName)
                sortKey = _sortList.ElementAt(ix).Key
                _sortList.Remove(sortKey)
            End If

            ' nur wenn es jetzt noch nicht drin ist, reintun .... 
            If Not _sortList.ContainsValue(cItem.projectName) Then
                ' jetzt das hproj bestimmen 
                Dim hproj As clsProjekt = AlleProjekte.getProject(key)
                If Not IsNothing(hproj) Then
                    sortKey = hproj.getSortKeyForConstellation(_sortType)

                    If Not _sortList.ContainsKey(sortKey) Then
                        _sortList.Add(sortKey, hproj.name)
                    Else
                        ' es muss ein . ergänzt werden 
                        sortKey = sortKey & "."
                        Do While _sortList.ContainsKey(sortKey)
                            sortKey = sortKey & "."
                        Loop
                        _sortList.Add(sortKey, hproj.name)
                    End If

                End If
            End If
        End If

    End Sub


    ''' <summary>
    ''' löscht den Eintrag mit Schlüssel key; wenn der nicht vorhanden ist, dann passiert gar nichts 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <remarks></remarks>
    Public Sub remove(key As String)

        If _allItems.ContainsKey(key) Then
            Dim cItem As clsConstellationItem = _allItems.Item(key)
            Dim pName As String = cItem.projectName

            _allItems.Remove(key)

            ' jetzt in der Sortliste entsprechend löschen und neu bestimmen , falls es in der Constellation 
            ' noch eine Variante des Projektes gibt ... 
            If _sortList.ContainsValue(pName) Then
                _sortList.RemoveAt(_sortList.IndexOfValue(pName))
            End If

            If Me.containsProject(pName) Then
                ' es gibt immer noch Varianten von pName 
                ' also neu bestimmen 
                Dim hproj As clsProjekt = Me.getSortRelevantProject(pName)
                If Not IsNothing(hproj) Then
                    Dim sortKey As String = hproj.getSortKeyForConstellation(_sortType)

                    If Not _sortList.ContainsKey(sortKey) Then
                        _sortList.Add(sortKey, hproj.name)
                    Else
                        ' es muss ein . ergänzt werden 
                        sortKey = sortKey & "."
                        Do While _sortList.ContainsKey(sortKey)
                            sortKey = sortKey & "."
                        Loop
                        _sortList.Add(sortKey, hproj.name)
                    End If

                End If
            Else
                ' wenn es keine Einträge mehr von pName gibt, dann muss es ggf aus der customLastList raus 
                If Not IsNothing(_lastCustomList) Then
                    If _lastCustomList.ContainsValue(pName) Then
                        _lastCustomList.RemoveAt(_lastCustomList.IndexOfValue(pName))
                    End If
                End If
                
            End If

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

        If Me._allItems.ContainsKey(pvName) Then

            Dim cItem As clsConstellationItem = Me._allItems.Item(pvName)
            If withShowFlag Then
                found = cItem.show
            Else
                found = True
            End If
        Else
            found = False
        End If

        contains = found
    End Function

    ''' <summary>
    ''' liefert true, wenn das Projekt in irgendeiner Form , mit oder ohne Varianten-NAme vorkommt
    ''' false, andernfalls
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsProject(ByVal pName As String) As Boolean
        Get
            Dim found As Boolean = False
            Dim index As Integer = 0

            Do While Not found And index <= _allItems.Count - 1
                Dim tmpName As String = getPnameFromKey(_allItems.ElementAt(index).Key)
                If tmpName = pName Then
                    found = True
                Else
                    index = index + 1
                End If
            Loop

            containsProject = found

        End Get
    End Property

    ''' <summary>
    ''' ähnlich wie reduceToElementsWithShow, aber hier werden nur die Projekte rausgeschmissen, die gar nicht in ShowProjekte sind bzw. die in ShowProjekte sind 
    ''' </summary>
    ''' <param name="requiredShowAttribute"></param>
    ''' <remarks></remarks>
    Public Sub reduceToProjectsWith(ByVal requiredShowAttribute As Boolean)
        Dim toDelete As New Collection

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In Me._allItems

            If requiredShowAttribute = ShowProjekte.contains(kvp.Value.projectName) Then
                ' nichts tun, soll ja nicht aus der Collection fliegen ...
            Else
                If Not toDelete.Contains(kvp.Key) Then
                    toDelete.Add(kvp.Key, kvp.Key)
                End If
            End If

        Next

        ' jetzt alle Einträge, die nicht in das Raster fallen, aus der Constellation löschen 
        For Each tmpName As String In toDelete

            If Me._allItems.ContainsKey(tmpName) Then
                ' da jetzt auch sort upgedated werden muss, die MEthode me.remove aufrufen
                'Me._allItems.Remove(tmpName)
                Me.remove(tmpName)
            End If

        Next

    End Sub
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
                ' da jetzt auch sort upgedated werden muss, die MEthode me.remove aufrufen
                'Me._allItems.Remove(tmpName)
                Me.remove(tmpName)
            End If

        Next

    End Sub

    ''' <summary>
    ''' ändert die Referenzen, die bisher auf oldvName gingen auf newVname 
    ''' wenn oldkey existiert, wird einfach der newKey in der Constellation gelöscht 
    ''' das ShowAttribute von pName (oldvName) muss übernommen werden ! 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="oldvName"></param>
    ''' <param name="newvName"></param>
    ''' <remarks></remarks>
    Public Sub updateVariantName(ByVal pName As String, ByVal oldvName As String, ByVal newvName As String)

        If oldvName = newvName Then
            ' nichts tun 
        Else
            ' da der pname unverändert bleibt, muss in _sortlist nichts getan werden ... 
            Dim oldKey As String = calcProjektKey(pName, oldvName)
            Dim newKey As String = calcProjektKey(pName, newvName)

            If _allItems.ContainsKey(oldKey) Then

                Dim cItem As clsConstellationItem = _allItems.Item(oldKey)

                ' das alte rausnehmen 
                _allItems.Remove(oldKey)

                ' umbenennen
                cItem.variantName = newvName

                ' in der Liste der  Items aufnehmen 
                ' wenn der schon existiert , rausnehmen ... und durch das mit dem Varianten Namen aktualsierte oldkey ersetzen 
                If _allItems.ContainsKey(newKey) Then
                    _allItems.Remove(newKey)
                End If
                _allItems.Add(newKey, cItem)

            End If
        End If




    End Sub
    ''' <summary>
    ''' sorgt dafür , dass in der Konstellation alle Projekte mit Name oldNAme mit dem neuen Namen bezeichnet werden 
    ''' </summary>
    ''' <param name="oldPName"></param>
    ''' <param name="newPname"></param>
    ''' <remarks></remarks>
    Public Function renameProject(ByVal oldPName As String, ByVal newPname As String) As Integer

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
            '_allItems.Remove(tmpName)
            Me.remove(tmpName)
        Next

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In toAddItems
            '_allItems.Add(kvp.Key, kvp.Value)
            Me.add(kvp.Value)
        Next

        renameProject = toAddItems.Count

    End Function

    Sub New()

        _allItems = New SortedList(Of String, clsConstellationItem)
        _sortList = New SortedList(Of String, String)
        _lastCustomList = Nothing
        _sortType = -1
        Me.constellationName = "" ' damit wird der Name Last<userName>

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
        _sortList = New SortedList(Of String, String)
        _lastCustomList = Nothing
        _sortType = -1

        Me.constellationName = cName

        If IsNothing(projektListe) Then
            ' bereits fertig - es ist eine leere Constellation mit Name cNAme
        Else

            If Not IsNothing(fullProjectNames) Then

                Dim newConstellationItem As clsConstellationItem
                _sortType = ptSortCriteria.alphabet

                For Each kvp As KeyValuePair(Of String, String) In fullProjectNames

                    Dim fullName As String = kvp.Key
                    Dim hproj As clsProjekt = projektListe.getProject(fullName)

                    If Not IsNothing(hproj) Then
                        newConstellationItem = New clsConstellationItem

                        With newConstellationItem
                            .projectName = hproj.name
                            .variantName = hproj.variantName
                            .zeile = 0
                            .start = hproj.startDate

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
                        If takeWhat = ptSzenarioConsider.all Or _
                            (takeWhat = ptSzenarioConsider.show And newConstellationItem.show) Or _
                            (takeWhat = ptSzenarioConsider.noshow And Not newConstellationItem.show) Then

                            Me.add(newConstellationItem)


                        End If


                    End If

                Next

            Else
                _sortType = ptSortCriteria.alphabet

                For Each kvp As KeyValuePair(Of String, clsProjekt) In projektListe.liste

                    Dim newConstellationItem As clsConstellationItem = New clsConstellationItem

                    With newConstellationItem
                        .projectName = kvp.Value.name
                        .variantName = kvp.Value.variantName
                        .zeile = 0
                        .start = kvp.Value.startDate

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
                    If takeWhat = ptSzenarioConsider.all Or _
                        (takeWhat = ptSzenarioConsider.show And newConstellationItem.show) Or _
                        (takeWhat = ptSzenarioConsider.noshow And Not newConstellationItem.show) Then

                        Me.add(newConstellationItem)

                    End If

                Next
            End If

        End If

    End Sub

End Class
