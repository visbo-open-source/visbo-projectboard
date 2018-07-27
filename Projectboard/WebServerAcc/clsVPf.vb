
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

    Public Property allItems As List(Of clsWebVPfItem)


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
        _allItems = New List(Of clsWebVPfItem)
    End Sub

    Sub copyfrom(ByVal c As clsConstellation)

        Dim sortElem As String = ""

        Me.name = c.constellationName

        For Each item In c.Liste
            Dim newItem As New clsWebVPfItem
            newItem.copyfrom(item.Value)
            newItem.name = newItem.projectName
            newItem.vpid = ""
            Me.allItems.Add(newItem)
        Next

        ' jetzt muss die Sortier-Reihenfolge und der Sortier-Typ gespeichert werden 
        ' dabei wird auch der Me.sortType gesetzt  

        If c.sortCriteria >= 0 Then
            Me.sortType = c.sortCriteria
        Else
            Me.sortType = ptSortCriteria.alphabet
        End If

        ' Kopieren der Sort-Liste 
        If Not IsNothing(c.sortListe) Then
            For Each kvp As KeyValuePair(Of String, String) In c.sortListe
                '' ??? ur: 2018.07.27: wird für normale list(of string) nicht benötigt?
                ''sortElem = kvp.Key
                ''If sortElem.Contains(punktName) Then
                ''    sortElem = sortElem.Replace(punktName, punktNameDB)
                ''End If
                Me.sortList.Add(kvp.Value)
            Next
        End If
        '' ur: 2018.07.24: beim Rest-Server gibt es keine lastCustomList mehr
        ''If Not IsNothing(c.lastCustomList) Then
        ''    ' die lastCustomList kopieren 
        ''    For Each kvp As KeyValuePair(Of String, String) In c.lastCustomList
        ''        sortElem = kvp.Key
        ''        If sortElem.Contains(punktName) Then
        ''            sortElem = sortElem.Replace(punktName, punktNameDB)
        ''        End If
        ''        Me.lastCustomList.Add(sortElem, kvp.Value)
        ''    Next

        ''End If

    End Sub


    ''' <summary>
    ''' kopiert eine Konstellation aus der Datenbank (via RestServer) in eine Hauptspeicher-Konstellation
    ''' </summary>
    ''' <param name="c"></param>
    ''' <remarks></remarks>
    Sub copyto(ByRef c As clsConstellation)
        Dim key As String
        'Dim sortElem As String

        c.constellationName = Me.name
        c.sortCriteria = Me.sortType

        For Each item In Me.allItems
            Dim newItem As New clsConstellationItem
            item.copyto(newItem)
            key = calcProjektKey(newItem.projectName, newItem.variantName)
            'key = item.projectName & "#" & item.variantName
            If Not c.Liste.ContainsKey(key) Then
                c.Liste.Add(key, newItem)
            Else
                Dim a As Integer = 0
                'Call MsgBox("Fehler bei Aufbau Konstellation mit Elem: " & key)
            End If


        Next

        If Not IsNothing(Me.sortList) And Not IsNothing(Me.sortType) Then
            If Me.sortList.Count = 0 And Me.allItems.Count > 0 Then
                ' mit diesem Befehl wird, wenn die SortListe Null ist, dieselbe auch gleich aufgebaut 

                ' in diesem Fall handelt es sich um eine "alte" Konstellation, die noch keine Sortliste enthält
                ' deshalb soll hier die Sortier-Reihenfolge gemäß tfzeile errechnet werden  
                Call c.buildSortlist(ptSortCriteria.customTF)

            Else
                ' hier wird die existierende Liste übernommen 
                If Not IsNothing(Me.sortList) Then

                    'c.sortListe(Me.sortType) = Me.sortList
                    For Each kvp As String In Me.sortList
                        '' ??? ur: 2018.07.27: wird für normale list(of string) nicht benötigt?
                        ''sortElem = kvp.Key
                        ''If sortElem.Contains(punktNameDB) Then
                        ''    sortElem = sortElem.Replace(punktNameDB, punktName)
                        ''End If
                        Call c.buildSortlist(Me.sortType)
                        ' ur:2018.07.27: zuvor: c.sortListe(Me.sortType).Add(sortElem, kvp.Value)
                    Next
                End If

            End If

            '' ur: 2018.07.24: lastCustomList gibt es beim Server nicht mehr
            ''If Not IsNothing(Me.lastCustomList) Then
            ''    ' die lastCustomList kopieren 

            ''    For Each kvp As KeyValuePair(Of String, String) In Me.lastCustomList
            ''        'c.lastCustomList.Add(kvp.Key, kvp.Value)
            ''        sortElem = kvp.Key
            ''        If sortElem.Contains(punktNameDB) Then
            ''            sortElem = sortElem.Replace(punktNameDB, punktName)
            ''        End If
            ''        c.lastCustomList.Add(sortElem, kvp.Value)
            ''    Next

            ''End If
        Else
            ' sorttype wird per Default auf alphabetisch sortiert setzen 
            ' damit wird auch _sortlist gesetzt ...
            c.sortCriteria = ptSortCriteria.alphabet

        End If


    End Sub
End Class
