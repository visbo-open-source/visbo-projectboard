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
            Dim deductBecausePoint As Integer = 0

            ' der wievielte Eintrag mit Attribut = Show ist es in der Liste ? 
            Do While ix <= _sortList.Count - 1 And Not found
                Dim vglName As String = _sortList.ElementAt(ix).Value
                If vglName = pName Then

                    If _sortType = ptSortCriteria.customTF Then
                        If _sortList.ElementAt(ix).Key.EndsWith(".") Then
                            deductBecausePoint = deductBecausePoint + 1
                        End If
                    End If

                    found = True
                Else

                    If Me.isShown(vglName) = True Then
                        bzeile = bzeile + 1

                        If _sortType = ptSortCriteria.customTF Then
                            If _sortList.ElementAt(ix).Key.EndsWith(".") Then
                                deductBecausePoint = deductBecausePoint + 1
                            End If
                        End If

                    End If

                    ix = ix + 1

                End If

            Loop

            If _sortType = ptSortCriteria.customTF Then
                bzeile = bzeile - deductBecausePoint
            End If
            

            getBoardZeile = bzeile + 2

        End Get
    End Property

    Private ReadOnly Property isShown(ByVal pName As String) As Boolean
        Get
            Dim ix As Integer = 0
            Dim found As Boolean = False

            Do While ix <= _allItems.Count - 1 And Not found
                If _allItems.ElementAt(ix).Value.projectName = pName And _
                    _allItems.ElementAt(ix).Value.show = True Then
                    found = True
                Else
                    ix = ix + 1
                End If
            Loop

            isShown = found
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
                        If Me.containsProject(kvp.Value) Then
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

    Public Sub buildSortlist(ByVal sCriteria As Integer)

        Dim key As String = ""

        ' die customTF Liste merken, wenn es sich darum gehandelt hat ... 
        If _sortType = ptSortCriteria.customTF Then
            _lastCustomList = _sortList
        End If

        ' jetzt müssen die Sort-Keys gesetzt werden 
        _sortType = sCriteria
        _sortList = New SortedList(Of String, String)

        ' das Folgende muss nur gemacht werden, wenn in AlleProjekte schon was drin ist 
        If sCriteria = ptSortCriteria.alphabet Then
            ' kann auch ohne AlleProjekte gemacht werden ... 
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                key = kvp.Value.projectName
                If Not _sortList.ContainsKey(key) Then
                    ' aufnehmen ...
                    _sortList.Add(key, key)
                Else
                    ' wenn es schon drin ist, muss nichts weiter gemacht werden 
                End If
            Next

        ElseIf sCriteria = ptSortCriteria.customTF Then

            ' neu 
            Dim newSortList As New SortedList(Of String, String)
            Dim noShowList As New SortedList(Of String, clsConstellationItem)

            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                ' erstmal prüfen , ob die sortliste das Projekt nicht schon enthält ...
                If kvp.Value.show = True Then
                    Dim sortkey As String = calcSortKeyCustomTF(kvp.Value.zeile)
                    ' jetzt wird der Schlüssel solange verändert, bis er eindeutig ist ... 
                    While newSortList.ContainsKey(sortkey)
                        sortkey = calcSortKeyCustomTF1(sortkey)
                    End While

                    ' jetzt ist er eindeutig 
                    newSortList.Add(sortkey, kvp.Value.projectName)
                Else
                    ' erstmal in die NoShow Liste packen 
                    noShowList.Add(calcProjektKey(kvp.Value.projectName, kvp.Value.variantName), kvp.Value)
                End If

            Next

            ' jetzt müssen alle NoShow-Items behandelt werden ..
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In noShowList
                If newSortList.ContainsValue(kvp.Value.projectName) Then
                    ' ist schon enthalten, also cItem.zeile anpassen 
                    Me.getItem(kvp.Key).zeile = getTFzeilefromSortKeyCustomTF _
                        (newSortList.ElementAt(newSortList.IndexOfValue(kvp.Value.projectName)).Key)
                Else
                    ' ist noch nicht enthalten, also ist das Projekt in keiner Variante angezeigt
                    ' und soll demzufolge eine Zeile-Nummer höher, also ans Ende positioniert werden 
                    Dim noShowZeile As Integer
                    If kvp.Value.zeile >= 2 Then
                        noShowZeile = kvp.Value.zeile
                    Else
                        If newSortList.Count > 0 Then
                            noShowZeile = getTFzeilefromSortKeyCustomTF _
                                                   (newSortList.Last.Key) + 1
                        Else
                            noShowZeile = 2
                        End If
                    End If

                    Dim tmpKey As String = calcSortKeyCustomTF(noShowZeile)
                    ' jetzt wird der Schlüssel solange verändert, bis er eindeutig ist ... 
                    While newSortList.ContainsKey(tmpKey)
                        tmpKey = calcSortKeyCustomTF1(tmpKey)
                    End While

                    ' jetzt ist er eindeutig 
                    newSortList.Add(tmpKey, kvp.Value.projectName)
                    Me.getItem(kvp.Key).zeile = noShowZeile

                End If
            Next

            ' jetzt enthält die newSortList alle Projekt-Namen mit den richtigen sortkeys ...
            Me.sortListe(ptSortCriteria.customTF) = newSortList


        ElseIf AlleProjekte.Count > 0 Then
            ' es handelt sich nicht um alphabet, nicht um CustomTF

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
                                ' es muss ein "x" ergänzt werden 
                                Do While _sortList.ContainsKey(key)
                                    key = calcSortKeyCustomTF1(key)
                                Loop
                                _sortList.Add(key, hproj.name)
                            End If
                        End If
                    End If


                End If

            Next
        End If

    End Sub

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
                    Call Me.buildSortlist(value)
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
                    _constellationName = value.Trim
                Else
                    _constellationName = calcLastSessionScenarioName()
                End If
            Else
                _constellationName = calcLastSessionScenarioName()
            End If
        End Set
    End Property

    Public Sub checkAndCorrectYourself(ByVal aktionskennung As Integer)

        If aktionskennung = PTTvActions.chgInSession Then

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
                    If ShowProjekte.contains(kvp.Value.projectName) Then

                        Dim hproj As clsProjekt = ShowProjekte.getProject(kvp.Value.projectName)
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

        End If
        

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
    ''' by default: Names coming from cItem-Liste
    ''' by default: independent of having show-Attribute or not
    ''' when considerShowAttr = true , only names with show-attribute = showvalue are in the output list 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getProjectNames(Optional ByVal fromCItemList As Boolean = True, _
                                             Optional ByVal considerShowAttribute As Boolean = False, _
                                             Optional ByVal showAttribute As Boolean = True) As SortedList(Of String, String)
        Get
            Dim tmpList As New SortedList(Of String, String)
            Dim pName As String

            If fromCItemList Then
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                    pName = kvp.Value.projectName

                    If considerShowAttribute Then
                        If kvp.Value.show = showAttribute Then
                            If Not tmpList.ContainsKey(pName) Then
                                tmpList.Add(key:=pName, value:=pName)
                            End If
                        End If
                    Else
                        If Not tmpList.ContainsKey(pName) Then
                            tmpList.Add(key:=pName, value:=pName)
                        End If
                    End If


                Next
            Else

                tmpList = _sortList

            End If

            getProjectNames = tmpList

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
    ''' setzt das oder die ShowAttribute gemäß der Variable showAttribute 
    ''' stellt sicher, dass in einer Constellation ein Projekt in max einer Variante das Attribut show haben kann 
    '''   
    ''' </summary>
    ''' <param name="pName">Projektname</param>
    ''' <param name="vName" >Varianten-Name </param>
    ''' <param name="showAttribute" >show: true; noShow:false</param>
    ''' <remarks></remarks>
    Public Sub updateShowAttributes(ByVal pName As String, ByVal vName As String, _
                                    ByVal showAttribute As Boolean)
        Dim currentProjectName As String = ""

        ' es werden alle Einträge gemäß Status Showprojekte aktualisiert 
        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
            ' den einen Namen behandeln 
            If pName = kvp.Value.projectName Then

                If showAttribute = True Then
                    If vName = kvp.Value.variantName Then
                        kvp.Value.show = True
                    Else
                        kvp.Value.show = False
                    End If
                ElseIf IsNothing(vName) Then
                    kvp.Value.show = False
                ElseIf vName = kvp.Value.variantName Then
                    kvp.Value.show = False
                End If

            End If

        Next


    End Sub


    ''' <summary>
    ''' kopiert eine Constellation, d.h jetzt müssen auch sortType und sortList kopiert werden
    ''' wird kein Name übergeben, wird der Name der zu kopierenden Constellation verwendet 
    ''' wird Last angegeben , so wird Last (username) verwendet  
    ''' </summary>
    ''' <param name="cName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property copy(Optional ByVal cName As String = "") As clsConstellation
        Get
            Dim copyResult As New clsConstellation

            ' wenn leer, soll der Name der zu kopierenden Konstellation verwendet werden 
            If cName = "" Then
                cName = Me.constellationName
            End If



            With copyResult
                .constellationName = cName

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In _allItems
                    Dim copiedItem As clsConstellationItem = kvp.Value.copy
                    .add(cItem:=copiedItem, noUpdateSortlist:=True)
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
    ''' setzt das cItem mit dem angegebenen Key auf citem.show = true 
    ''' stellt sicher, dass alle anderen Items auf noShow gesetzt werden 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="showAttribute"></param>
    ''' <remarks></remarks>
    Public Sub setItemToShow(ByVal key As String, ByVal showAttribute As Boolean)

        If _allItems.ContainsKey(key) Then

            Dim pName As String = getPnameFromKey(key)
            Dim vName As String = getVariantnameFromKey(key)
            If showAttribute = True Then
                Call Me.setToNoShowExcept(pName, vName)
            End If

            _allItems.Item(key).show = showAttribute
        End If

    End Sub
    ''' <summary>
    ''' fügt ein clsConstellationItem hinzu und aktualisiert auch die Sortlist entsprechend ... 
    ''' Voraussetzung: in AlleProjekte ist das im Item beschriebene Objekt bereits enthalten 
    ''' im add muss kein Update der lastCustomlist erfolgen, nur beim Remove ... 
    ''' </summary>
    ''' <param name="cItem"></param>
    ''' <remarks></remarks>
    Public Sub add(ByVal cItem As clsConstellationItem, _
                   Optional ByVal sKey As Integer = -1, _
                   Optional ByVal noUpdateSortlist As Boolean = False)

        Dim key As String
        Dim sortKey As String = ""
        'key = cItem.projectName & "#" & cItem.variantName
        key = calcProjektKey(cItem.projectName, cItem.variantName)

        ' wenn cItem.show = true, dann alle anderen von diesem Projekt auf noShow setzen 
        If cItem.show = True Then
            Call Me.setToNoShowExcept(cItem.projectName, cItem.variantName)
        End If

        ' wenn jetzt der Schlüssel bereits vorkommt, dann den Schlüssel löschen 
        If _allItems.ContainsKey(key) Then
            _allItems.Remove(key)
        End If

        ' jetzt wird der Schlüssel aufgenommen 
        _allItems.Add(key, cItem)

        ' soll die Sortliste upgedated werden 
        ' gibt es bereits einen Eintrag in der _sortliste ? 
        If Not noUpdateSortlist Then
            ' jetzt auch in sortlist aktualisieren 
            ' der alte Schlüssel soll nur dann rausgenommen werden, 
            ' wenn er aufgrund einer neuen Variante neu berechnet werden muss 
            ' und nicht alphabetisch, Custom-Liste oder TF gesteuert ist 
            If _sortList.ContainsValue(cItem.projectName) And cItem.show And _
                Not (_sortType = ptSortCriteria.customTF Or _
                     _sortType = ptSortCriteria.customListe Or _
                     _sortType = ptSortCriteria.alphabet) Then
                ' Remove den bisherigen Schlüssel 
                Dim ix As Integer = _sortList.IndexOfValue(cItem.projectName)
                sortKey = _sortList.ElementAt(ix).Key
                _sortList.Remove(sortKey)
            End If

            ' nur wenn es jetzt noch nicht drin ist, reintun .... 
            If Not _sortList.ContainsValue(cItem.projectName) Then

                If _sortType = ptSortCriteria.alphabet Then
                    sortKey = cItem.projectName

                ElseIf _sortType = ptSortCriteria.customTF Then
                    Dim position As Integer
                    If sKey = -1 Then
                        position = _sortList.Count + 2
                    Else
                        position = sKey
                    End If

                    sortKey = calcSortKeyCustomTF(position)

                Else
                    ' jetzt das hproj bestimmen 
                    Dim hproj As clsProjekt = AlleProjekte.getProject(key)
                    If Not IsNothing(hproj) Then
                        sortKey = hproj.getSortKeyForConstellation(_sortType)
                    End If
                End If

                If Not _sortList.ContainsKey(sortKey) Then
                    _sortList.Add(sortKey, cItem.projectName)
                Else
                    ' es muss ein . ergänzt werden 
                    Do While _sortList.ContainsKey(sortKey)
                        sortKey = calcSortKeyCustomTF1(key)
                    Loop
                    _sortList.Add(sortKey, cItem.projectName)
                End If


            End If
        End If


    End Sub

    ''' <summary>
    ''' setzt in der angegebenen Constellation alle items to Nowshow, ausser den angegebenen Werte-Paar 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <remarks></remarks>
    Private Sub setToNoShowExcept(ByVal pName As String, ByVal vName As String)
        Dim found As Boolean = False
        Dim finished As Boolean = False
        Dim ix As Integer = 0
        Dim anzItems As Integer = _allItems.Count

        If anzItems = 0 Then
            ' nichts zu tun 
        Else
            Do While Not found And ix <= anzItems - 1
                If _allItems.ElementAt(ix).Value.projectName <> pName Then
                    ix = ix + 1
                Else
                    found = True
                End If

            Loop

            If ix > anzItems - 1 Then
                ' nichts tun, fertig 
            Else
                finished = False
                Do While Not finished And ix <= anzItems - 1
                    If _allItems.ElementAt(ix).Value.projectName = pName Then
                        If _allItems.ElementAt(ix).Value.variantName <> vName Then
                            _allItems.ElementAt(ix).Value.show = False
                        End If
                        ix = ix + 1
                    Else
                        finished = True
                    End If
                Loop

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
                        Do While _sortList.ContainsKey(sortKey)
                            sortKey = calcSortKeyCustomTF1(sortKey)
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

    Public Sub updateTFzeile(ByVal key As String, ByVal tfzeile As Integer)

        If _allItems.ContainsKey(key) Then
            _allItems.Item(key).zeile = tfzeile
        End If

        ' nur wenn der _sorttype = customTF ist, dann aktualisieren des Schlüssels 
        If _sortType = ptSortCriteria.customTF Then
            Dim pName As String = getPnameFromKey(key)
            If pName <> "" Then
                Dim ix As Integer = _sortList.IndexOfValue(pName)

                If ix >= 0 Then
                    ' den alten eintrag rausnehmen 
                    _sortList.RemoveAt(ix)

                    ' den neuen Schlüssel bestimmen  
                    Dim sortKey As String = calcSortKeyCustomTF(tfzeile)
                    While _sortList.ContainsKey(sortKey)
                        sortKey = calcSortKeyCustomTF1(sortKey)
                    End While

                    ' den neuen Schlüssel eintragen 
                    _sortList.Add(sortKey, pName)

                End If
            End If
            
        End If


    End Sub

    ''' <summary>
    ''' ändert die Referenzen, die bisher auf oldvName gingen auf newVname 
    ''' wenn oldkey existiert, wird einfach der oldkey in der Constellation gelöscht 
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
        Me.constellationName = "" ' damit wird der Name Last (<userName>)

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
            Optional ByVal cName As String = "", _
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
