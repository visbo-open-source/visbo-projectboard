Public Class clsConstellationDB

    Public allItems As List(Of clsConstellationItemDB)
    Public constellationName As String
    Public sortType As Integer
    Public sortList As SortedList(Of String, String)
    Public lastCustomList As SortedList(Of String, String)
    Public Id As String

    Sub copyfrom(ByVal c As clsConstellation)

        Dim sortElem As String = ""

        Me.constellationName = c.constellationName

        For Each item In c.Liste
            Dim newItem As New clsConstellationItemDB
            newItem.copyfrom(item.Value)
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
                sortElem = kvp.Key
                If sortElem.Contains(punktName) Then
                    sortElem = sortElem.Replace(punktName, punktNameDB)
                End If
                Me.sortList.Add(sortElem, kvp.Value)
            Next
        End If

        If Not IsNothing(c.lastCustomList) Then
            ' die lastCustomList kopieren 
            For Each kvp As KeyValuePair(Of String, String) In c.lastCustomList
                sortElem = kvp.Key
                If sortElem.Contains(punktName) Then
                    sortElem = sortElem.Replace(punktName, punktNameDB)
                End If
                Me.lastCustomList.Add(sortElem, kvp.Value)
            Next

        End If

    End Sub

    ''' <summary>
    ''' kopiert eine Konstellation aus der Datenbank in eine Hauptspeicher-Konstellation
    ''' </summary>
    ''' <param name="c"></param>
    ''' <remarks></remarks>
    Sub copyto(ByRef c As clsConstellation)
        Dim key As String
        Dim sortElem As String

        c.constellationName = Me.constellationName

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
                ' mit diesem Befehlt wird, wenn die SortListe Null ist, dieselbe auch gleich aufgebaut 

                ' in diesem Fall handelt es sich um eine "alte" Konstellation, die noch keine Sortliste enthält
                ' deshalb soll hier die Sortier-Reihenfolge gemäß tfzeile errechnet werden  
                Call c.buildSortlist(ptSortCriteria.customTF)

            Else
                ' hier wird die existierende Liste übernommen 
                If Not IsNothing(Me.sortList) Then
                    
                    'c.sortListe(Me.sortType) = Me.sortList
                    For Each kvp As KeyValuePair(Of String, String) In Me.sortList
                        sortElem = kvp.Key
                        If sortElem.Contains(punktNameDB) Then
                            sortElem = sortElem.Replace(punktNameDB, punktName)
                        End If
                        c.sortListe(Me.sortType).Add(sortElem, kvp.Value)
                    Next
                End If

            End If

            If Not IsNothing(Me.lastCustomList) Then
                ' die lastCustomList kopieren 

                For Each kvp As KeyValuePair(Of String, String) In Me.lastCustomList
                    'c.lastCustomList.Add(kvp.Key, kvp.Value)
                    sortElem = kvp.Key
                    If sortElem.Contains(punktNameDB) Then
                        sortElem = sortElem.Replace(punktNameDB, punktName)
                    End If
                    c.lastCustomList.Add(sortElem, kvp.Value)
                Next

            End If
        Else
            ' sorttype wird per Default auf alphabetisch sortiert setzen 
            ' damit wird auch _sortlist gesetzt ...
            c.sortCriteria = ptSortCriteria.alphabet

        End If


    End Sub

    Public Class clsConstellationItemDB
        Public projectName As String = ""
        Public variantName As String = ""
        Public Start As Date = StartofCalendar.AddMonths(-1)
        Public show As Boolean = True
        Public zeile As Integer = 0
        ' warum wird die entsprechende Projektvariante aufgenommen 
        Public reasonToInclude As String = ""
        ' warum wird die entsprechende Projektvariante nicht aufgenommen 
        Public reasonToExclude As String = ""

        Sub copyfrom(ByVal item As clsConstellationItem)

            With item
                Me.projectName = .projectName
                Me.variantName = .variantName
                Me.Start = .start.ToUniversalTime
                Me.show = .show
                Me.zeile = .zeile
                Me.reasonToInclude = .reasonToInclude
                Me.reasonToExclude = .reasonToExclude
            End With
        End Sub

        Sub copyto(ByRef item As clsConstellationItem)

            With item
                .projectName = Me.projectName
                .variantName = Me.variantName
                .start = Me.Start.ToLocalTime
                .show = Me.show
                .zeile = Me.zeile
                .reasonToInclude = Me.reasonToInclude
                .reasonToExclude = Me.reasonToExclude
            End With

        End Sub

        Sub New()
            
        End Sub

    End Class

    Sub New()

        allItems = New List(Of clsConstellationItemDB)
        sortType = 0
        sortList = New SortedList(Of String, String)
        lastCustomList = New SortedList(Of String, String)
    End Sub

End Class
