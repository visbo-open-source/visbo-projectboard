Public Class clsCapas

    Private _liste As List(Of clsCapa)
    Public ReadOnly Property liste As List(Of clsCapa)
        Get
            liste = _liste
        End Get
    End Property

    ''' <summary>
    ''' ensures that a tuple roleid, startofYear.year does only exist once in the list
    ''' adds the capa item if it does not yet exist. Existing means having the same roleID and the same starofYear
    ''' if it does already exist then the values are replaced 
    ''' </summary>
    ''' <param name="capa"></param>
    Public Sub Add(ByVal capa As clsCapa)
        Dim found As Boolean = False
        Dim foundItem As clsCapa = Nothing

        For Each capaItem As clsCapa In _liste
            found = (capaItem.roleID = capa.roleID) And
                    (capaItem.startOfYear.Year = capa.startOfYear.Year)

            If found Then
                foundItem = capaItem
            End If
        Next

        If found Then
            foundItem.capaPerMonth = capa.capaPerMonth
        Else
            _liste.Add(capa)
        End If

    End Sub

    Public Function Remove(ByVal uid As String) As clsCapas
        Dim tmpResult As New clsCapas

        For Each capaItem As clsCapa In _liste
            If capaItem.roleID <> uid Then
                tmpResult.Add(capaItem)
            End If
        Next

        Remove = tmpResult
    End Function

    Public Function Count(ByVal uid As String) As Integer
        Dim tmpResult As Integer = 0
        For Each capaItem As clsCapa In _liste
            If capaItem.roleID = uid Then
                tmpResult = tmpResult + 1
            End If
        Next
        Count = tmpResult
    End Function

    Public Function Remove(ByVal uid As String, ByVal year As Integer) As clsCapas
        Dim tmpResult As New clsCapas

        For Each capaItem As clsCapa In _liste
            If (capaItem.roleID <> uid) And (capaItem.startOfYear.Year <> year) Then
                tmpResult.Add(capaItem)
            End If
        Next

        Remove = tmpResult
    End Function

    Public Function containsIdentical(ByVal subset As clsCapas) As Boolean

        Dim myUid As String = ""
        If subset.liste.Count = 0 Then
            ' ist falsch 
        Else
            myUid = subset.liste.First.roleID
        End If

        Dim isIdentical As Boolean = Count(myUid) = subset.Count(myUid)

        If isIdentical Then
            For Each subsetItem As clsCapa In subset.liste

                Dim atleastOne As Boolean = False
                For Each capaItem As clsCapa In _liste
                    atleastOne = atleastOne Or subsetItem.isIdenticalTo(capaItem)
                Next

                isIdentical = isIdentical And atleastOne

                If Not isIdentical Then
                    Exit For
                End If

            Next
        End If

        containsIdentical = isIdentical

    End Function
    Public Sub New()
        _liste = New List(Of clsCapa)
    End Sub
End Class
