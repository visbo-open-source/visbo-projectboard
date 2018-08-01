
Imports ProjectBoardDefinitions
Public Class clsVPf
    Public Property _id As String
    Public Property name As String
    Public Property vpid As String
    Public Property variantName As String
    Public Property timestamp As String
    Public Property updatedAt As String
    Public Property createdAt As String
    Public Property sortType As Integer
    Public Property sortList As List(Of String)

    Public Property allItems As List(Of clsVPfItem)


    Sub New()
        _id = ""
        _name = "not named"
        _vpid = "not yet defined"
        _variantName = ""
        _timestamp = Date.MinValue.ToString
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
        _sortType = 1
        _sortList = New List(Of String)
        _allItems = New List(Of clsVPfItem)
    End Sub

    'Sub copyfrom(ByVal c As clsConstellation)

    '    Dim sortElem As String = ""

    '    Me.name = c.constellationName

    '    For Each item As KeyValuePair(Of String, clsConstellationItem) In c.Liste
    '        Dim newItem As New clsVPfItem
    '        newItem.copyfrom(item.Value)
    '        newItem.name = newItem.projectName
    '        newItem.vpid = ""
    '        Me.allItems.Add(newItem)
    '    Next

    '    ' jetzt muss die Sortier-Reihenfolge und der Sortier-Typ gespeichert werden 
    '    ' dabei wird auch der Me.sortType gesetzt  

    '    If c.sortCriteria >= 0 Then
    '        Me.sortType = c.sortCriteria
    '    Else
    '        Me.sortType = ptSortCriteria.alphabet
    '    End If

    '    ' Kopieren der Sort-Liste 
    '    If Not IsNothing(c.sortListe) Then
    '        For Each kvp As KeyValuePair(Of String, String) In c.sortListe
    '            '' ??? ur: 2018.07.27: wird für normale list(of string) nicht benötigt?
    '            ''sortElem = kvp.Key
    '            ''If sortElem.Contains(punktName) Then
    '            ''    sortElem = sortElem.Replace(punktName, punktNameDB)
    '            ''End If
    '            Me.sortList.Add(kvp.Value)
    '        Next
    '    End If
    '    '' ur: 2018.07.24: beim Rest-Server gibt es keine lastCustomList mehr
    '    ''If Not IsNothing(c.lastCustomList) Then
    '    ''    ' die lastCustomList kopieren 
    '    ''    For Each kvp As KeyValuePair(Of String, String) In c.lastCustomList
    '    ''        sortElem = kvp.Key
    '    ''        If sortElem.Contains(punktName) Then
    '    ''            sortElem = sortElem.Replace(punktName, punktNameDB)
    '    ''        End If
    '    ''        Me.lastCustomList.Add(sortElem, kvp.Value)
    '    ''    Next

    '    ''End If

    'End Sub


    '''' <summary>
    '''' kopiert eine Konstellation aus der Datenbank (via RestServer) in eine Hauptspeicher-Konstellation
    '''' </summary>
    '''' <param name="c"></param>
    '''' <remarks></remarks>
    'Sub copyto(ByRef c As clsConstellation)
    '    Dim key As String
    '    'Dim sortElem As String

    '    c.constellationName = Me.name
    '    c.sortCriteria = Me.sortType

    '    For Each item In Me.allItems
    '        Dim newItem As New clsConstellationItem
    '        item.copyto(newItem)
    '        key = calcProjektKey(newItem.projectName, newItem.variantName)
    '        'key = item.projectName & "#" & item.variantName
    '        If Not c.Liste.ContainsKey(key) Then
    '            c.Liste.Add(key, newItem)
    '        Else
    '            Dim a As Integer = 0
    '            'Call MsgBox("Fehler bei Aufbau Konstellation mit Elem: " & key)
    '        End If


    '    Next

    '    If Not IsNothing(Me.sortList) And Not IsNothing(Me.sortType) Then
    '        If Me.sortList.Count = 0 And Me.allItems.Count > 0 Then
    '            ' mit diesem Befehl wird, wenn die SortListe Null ist, dieselbe auch gleich aufgebaut 

    '            ' in diesem Fall handelt es sich um eine "alte" Konstellation, die noch keine Sortliste enthält
    '            ' deshalb soll hier die Sortier-Reihenfolge gemäß tfzeile errechnet werden  
    '            Call c.buildSortlist(Me.sortType)

    '        Else
    '            ' hier wird die existierende Liste übernommen 
    '            If Not IsNothing(Me.sortList) Then

    '                'c.sortListe(Me.sortType) = Me.sortList
    '                For Each kvp As String In Me.sortList
    '                    '' ??? ur: 2018.07.27: wird für normale list(of string) nicht benötigt?
    '                    ''sortElem = kvp.Key
    '                    ''If sortElem.Contains(punktNameDB) Then
    '                    ''    sortElem = sortElem.Replace(punktNameDB, punktName)
    '                    ''End If
    '                    Call c.buildSortlist(Me.sortType)
    '                    ' ur:2018.07.27: zuvor: c.sortListe(Me.sortType).Add(sortElem, kvp.Value)
    '                Next
    '            End If

    '        End If

    '        '' ur: 2018.07.24: lastCustomList gibt es beim Server nicht mehr
    '        ''If Not IsNothing(Me.lastCustomList) Then
    '        ''    ' die lastCustomList kopieren 

    '        ''    For Each kvp As KeyValuePair(Of String, String) In Me.lastCustomList
    '        ''        'c.lastCustomList.Add(kvp.Key, kvp.Value)
    '        ''        sortElem = kvp.Key
    '        ''        If sortElem.Contains(punktNameDB) Then
    '        ''            sortElem = sortElem.Replace(punktNameDB, punktName)
    '        ''        End If
    '        ''        c.lastCustomList.Add(sortElem, kvp.Value)
    '        ''    Next

    '        ''End If
    '    Else
    '        ' sorttype wird per Default auf alphabetisch sortiert setzen 
    '        ' damit wird auch _sortlist gesetzt ...
    '        c.sortCriteria = ptSortCriteria.alphabet

    '    End If


    'End Sub

    'Public Function buildSortlist(ByVal c As clsVPf, ByVal sCriteria As Integer) As SortedList(Of String, String)

    '    Dim key As String = ""
    '    Dim result_sortList As New SortedList(Of String, String)

    '    ' die customTF Liste merken, wenn es sich darum gehandelt hat ... 
    '    If c.sortType = ptSortCriteria.customTF And c.sortType <> sCriteria Then
    '        ' Kopieren der Liste 
    '        Dim lastCustomList As New SortedList(Of String, String)
    '        Dim lfdNr As Integer = 2
    '        For Each vpid As String In c.sortList

    '            Dim pname As String = GETpName(vpid)
    '            lastCustomList.Add(lfdNr.ToString, pname)

    '        Next

    '    End If

    '    ' jetzt müssen die Sort-Keys gesetzt werden 
    '    c.sortType = sCriteria


    '    ' das Folgende muss nur gemacht werden, wenn in AlleProjekte schon was drin ist 
    '    If sCriteria = ptSortCriteria.alphabet Then
    '        ' kann auch ohne AlleProjekte gemacht werden ... 
    '        For Each vpfItem As clsVPfItem In c.allItems
    '            key = vpfItem.projectName
    '            If Not result_sortList.ContainsKey(key) Then
    '                ' aufnehmen ...
    '                result_sortList.Add(key, key)
    '            Else
    '                ' wenn es schon drin ist, muss nichts weiter gemacht werden 
    '            End If
    '        Next

    '    ElseIf sCriteria = ptSortCriteria.customTF Then

    '        ' neu 
    '        Dim newSortList As New SortedList(Of String, String)
    '        Dim noShowList As New SortedList(Of String, clsVPfItem)

    '        For Each vpfItem As clsVPfItem In c.allItems
    '            ' erstmal prüfen , ob die sortliste das Projekt nicht schon enthält ...
    '            If vpfItem.show = True Then
    '                Dim sortkey As String = calcSortKeyCustomTF(vpfItem.zeile)
    '                ' jetzt wird der Schlüssel solange verändert, bis er eindeutig ist ... 
    '                While newSortList.ContainsKey(sortkey)
    '                    sortkey = calcSortKeyCustomTF1(sortkey)
    '                End While

    '                ' jetzt ist er eindeutig 
    '                newSortList.Add(sortkey, vpfItem.projectName)
    '            Else
    '                ' erstmal in die NoShow Liste packen 
    '                noShowList.Add(calcProjektKey(vpfItem.projectName, vpfItem.variantName), vpfItem)
    '            End If

    '        Next

    '        ' jetzt müssen alle NoShow-Items behandelt werden ..
    '        For Each kvp As KeyValuePair(Of String, clsVPfItem) In noShowList
    '            If newSortList.ContainsValue(kvp.Value.projectName) Then
    '                ' ist schon enthalten, also cItem.zeile anpassen 
    '                kvp.Value.zeile = getTFzeilefromSortKeyCustomTF _
    '                    (newSortList.ElementAt(newSortList.IndexOfValue(kvp.Value.projectName)).Key)
    '            Else
    '                ' ist noch nicht enthalten, also ist das Projekt in keiner Variante angezeigt
    '                ' und soll demzufolge eine Zeile-Nummer höher, also ans Ende positioniert werden 
    '                Dim noShowZeile As Integer
    '                If kvp.Value.zeile >= 2 Then
    '                    noShowZeile = kvp.Value.zeile
    '                Else
    '                    If newSortList.Count > 0 Then
    '                        noShowZeile = getTFzeilefromSortKeyCustomTF _
    '                                               (newSortList.Last.Key) + 1
    '                    Else
    '                        noShowZeile = 2
    '                    End If
    '                End If

    '                Dim tmpKey As String = calcSortKeyCustomTF(noShowZeile)
    '                ' jetzt wird der Schlüssel solange verändert, bis er eindeutig ist ... 
    '                While newSortList.ContainsKey(tmpKey)
    '                    tmpKey = calcSortKeyCustomTF1(tmpKey)
    '                End While

    '                ' jetzt ist er eindeutig 
    '                newSortList.Add(tmpKey, kvp.Value.projectName)
    '                kvp.Value.zeile = noShowZeile

    '            End If
    '        Next

    '        ' jetzt enthält die newSortList alle Projekt-Namen mit den richtigen sortkeys ...
    '        result_sortList = newSortList


    '    ElseIf AlleProjekte.Count > 0 Then
    '        ' es handelt sich nicht um alphabet, nicht um CustomTF

    '        For Each vpfItem As clsVPfItem In c.allItems

    '            Dim projkey As String = calcProjektKey(vpfItem.projectName, vpfItem.variantName)

    '            Dim hproj As clsProjekt = AlleProjekte.getProject(projkey)

    '            If Not IsNothing(hproj) Then

    '                ' nur wenn es nicht bereits in sortList enthalten ist 
    '                If Not result_sortList.ContainsValue(hproj.name) Then

    '                    hproj = getSortRelevantProject(c, hproj.name)

    '                    If Not IsNothing(hproj) Then
    '                        key = hproj.getSortKeyForConstellation(c.sortType)

    '                        If Not result_sortList.ContainsKey(key) Then
    '                            result_sortList.Add(key, hproj.name)
    '                        Else
    '                            ' es muss ein "x" ergänzt werden 
    '                            Do While result_sortList.ContainsKey(key)
    '                                key = calcSortKeyCustomTF1(key)
    '                            Loop
    '                            result_sortList.Add(key, hproj.name)
    '                        End If
    '                    End If
    '                End If


    '            End If

    '        Next
    '    End If

    '    buildSortlist = result_sortList

    'End Function

    '''' <summary>
    '''' provides the list of variant Names of an Portfolio c in alphabetical order 
    '''' if mitKlammer = true then items will enclosed by ()
    '''' </summary>
    '''' <param name="c"></param>
    '''' <param name="pName"></param>
    '''' <param name="mitKlammer"></param>
    '''' <returns></returns>
    'Private Function getVariantNames(ByVal c As clsVPf, ByVal pName As String, ByVal mitKlammer As Boolean) As Collection

    '    Dim tmpCollection As New Collection
    '    Dim vName As String

    '    For Each vpfItem As clsVPfItem In c.allItems

    '        If pName = vpfItem.projectName Then
    '            If mitKlammer Then
    '                vName = "(" & vpfItem.variantName & ")"
    '            Else
    '                vName = vpfItem.variantName
    '            End If

    '            tmpCollection.Add(vName)

    '        End If

    '    Next

    '    getVariantNames = tmpCollection

    'End Function
    '''' <summary>
    '''' liefert zu einem gegebenen Projekt-Namen das Projekt ab, das für die Sortier-Schlüssel-Berechnung verwendet werden soll 
    '''' das relevante Projekt ist das, was im Show ist bzw das was als erstes in der Variant-Liste steht  
    '''' Nothing, wenn es das Projekt gar nicht gibt 
    '''' </summary>
    '''' <param name="pName"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Private Function getSortRelevantProject(ByVal c As clsVPf, ByVal pName As String) As clsProjekt

    '    Dim hproj As clsProjekt = Nothing

    '    If ShowProjekte.contains(pName) Then
    '        hproj = ShowProjekte.getProject(pName)
    '    Else
    '        ' bestimme das hproj, das als erste Variante vorkommt 
    '        Dim vName As String = ""
    '        Dim tmpCollection As Collection = getVariantNames(c, pName, False)
    '        If Not IsNothing(tmpCollection) Then
    '            vName = CStr(tmpCollection.Item(1))
    '        End If
    '        Dim tmpKey As String = calcProjektKey(pName, vName)
    '        hproj = AlleProjekte.getProject(tmpKey)
    '    End If

    '    getSortRelevantProject = hproj

    'End Function

End Class
