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
