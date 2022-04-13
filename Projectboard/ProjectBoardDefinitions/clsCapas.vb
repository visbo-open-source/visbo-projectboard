Public Class clsCapas

    Private _liste As List(Of clsCapa)
    Public Property liste As List(Of clsCapa)
        Get
            liste = _liste
        End Get
        Set(value As List(Of clsCapa))
            If Not IsNothing(value) Then
                _liste = value
            Else
                _liste = New List(Of clsCapa)
            End If
        End Set
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
                    (capaItem.startOfYear.Year = capa.startOfYear.Year) And
                    (DateDiff(DateInterval.Day, capaItem.startOfYear, capa.startOfYear) = 0)

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

    Public Function Count(ByVal uid As String) As Integer
        Dim tmpResult As Integer = 0

        For Each capaItem As clsCapa In _liste
            If capaItem.roleID = uid Then
                tmpResult = tmpResult + 1
            End If
        Next
        Count = tmpResult
    End Function

    ''' <summary>
    ''' returns a clsCapa = List(of clsCapa) containing only items with regard to uid
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <returns></returns>
    Public Function getCapasOfOneRole(ByVal uid As String) As clsCapas
        Dim tmpResult As New clsCapas

        For Each capaItem As clsCapa In _liste
            If capaItem.roleID = uid Then
                tmpResult.Add(capaItem)
            End If
        Next

        getCapasOfOneRole = tmpResult
    End Function

    Public Function minus(ByVal myCapas As clsCapas) As clsCapas
        Dim tmpResult As New clsCapas

        For Each capaItem As clsCapa In _liste

            If IsNothing(myCapas.getCapa(capaItem.roleID, capaItem.startOfYear)) Then
                tmpResult.Add(capaItem)
            End If

        Next

        minus = tmpResult
    End Function

    Private Function getCapa(ByVal uid As String, ByVal myDate As Date) As clsCapa
        Dim tmpResult As clsCapa = Nothing

        For Each capaItem As clsCapa In _liste
            If (capaItem.roleID = uid) And (DateDiff(DateInterval.Day, capaItem.startOfYear, myDate) = 0) Then
                tmpResult = capaItem
                Exit For
            End If
        Next

        getCapa = tmpResult
    End Function

    Public Function Remove(ByVal uid As String) As clsCapas
        Dim tmpResult As New clsCapas

        For Each capaItem As clsCapa In _liste
            If capaItem.roleID <> uid Then
                tmpResult.Add(capaItem)
            End If
        Next

        Remove = tmpResult
    End Function

    Public Function Remove(ByVal uid As String, ByVal myDate As Date) As clsCapas
        Dim tmpResult As New clsCapas

        For Each capaItem As clsCapa In _liste
            If (capaItem.roleID <> uid) Or (DateDiff(DateInterval.Day, capaItem.startOfYear, myDate) <> 0) Then
                tmpResult.Add(capaItem)
            End If
        Next

        Remove = tmpResult
    End Function

    ''' <summary>
    ''' returns true if myCapa is exactly contained, i.e with respect to uid, startofYear, values of single months  
    ''' </summary>
    ''' <param name="myCapa"></param>
    ''' <returns></returns>
    Public Function containsIdentical(ByVal myCapa As clsCapa) As Boolean
        Dim result As Boolean = False

        If Not IsNothing(myCapa) Then
            Dim capaItem As clsCapa = getCapa(myCapa.roleID, myCapa.startOfYear)

            If Not IsNothing(capaItem) Then
                result = capaItem.isIdenticalTo(myCapa)
            End If
        End If

        containsIdentical = result
    End Function


    Public Function containsIdentical(ByVal subset As clsCapas) As Boolean

        Dim myUid As String = ""
        Dim isIdentical As Boolean = False

        If Not IsNothing(subset) Then

            If subset.liste.Count = 0 Then
                ' ist falsch 
            Else
                myUid = subset.liste.First.roleID
            End If

            isIdentical = (Count(myUid) = subset.Count(myUid))

            If isIdentical Then
                For Each subsetItem As clsCapa In subset.liste

                    Dim atleastOne As Boolean = False
                    For Each capaItem As clsCapa In _liste
                        atleastOne = atleastOne Or subsetItem.isIdenticalTo(capaItem)
                        If atleastOne = True Then
                            Exit For
                        End If
                    Next

                    isIdentical = isIdentical And atleastOne

                    If Not isIdentical Then
                        Exit For
                    End If

                Next
            End If
        End If


        containsIdentical = isIdentical

    End Function
    Public Sub New()
        _liste = New List(Of clsCapa)
    End Sub
End Class
