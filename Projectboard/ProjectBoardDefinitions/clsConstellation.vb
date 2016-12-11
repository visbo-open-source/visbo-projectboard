Public Class clsConstellation

    Private allItems As SortedList(Of String, clsConstellationItem)

    Public Property constellationName As String

    Public ReadOnly Property Liste() As SortedList(Of String, clsConstellationItem)

        Get
            Liste = allItems
        End Get

    End Property


    Public ReadOnly Property getItem(key As String) As clsConstellationItem

        Get
            getItem = allItems(key)
        End Get

    End Property

    Public ReadOnly Property count() As Integer

        Get
            count = allItems.Count
        End Get

    End Property

    Public Sub Add(cItem As clsConstellationItem)

        Dim key As String
        'key = cItem.projectName & "#" & cItem.variantName
        key = calcProjektKey(cItem.projectName, cItem.variantName)
        allItems.Add(key, cItem)

    End Sub


    Public Sub Remove(key As String)

        allItems.Remove(key)

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

        Do While ix <= Me.allItems.Count - 1 And Not found

            If pvName = Me.allItems.ElementAt(ix).Key Then
                If withShowFlag Then
                    If Me.allItems.ElementAt(ix).Value.show = True Then
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

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In Me.allItems
            If kvp.Value.show <> showAttribute Then
                If Not toDelete.Contains(kvp.Key) Then
                    toDelete.Add(kvp.Key, kvp.Key)
                End If

            End If
        Next

        ' jetzt alle Einträge, die nicht das showAttribute trugen, löschen 
        For Each tmpName As String In toDelete

            If Me.allItems.ContainsKey(tmpName) Then
                Me.allItems.Remove(tmpName)
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

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In allItems
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
            allItems.Remove(tmpName)
        Next

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In toAddItems
            allItems.Add(kvp.Key, kvp.Value)
        Next

        rename = toAddItems.Count

    End Function

    Sub New()

        allItems = New SortedList(Of String, clsConstellationItem)

    End Sub

End Class
