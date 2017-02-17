Public Class clsWriteProtections
    Private _allWriteProtections As SortedList(Of String, clsWriteProtectionItem)

    Public Property liste As SortedList(Of String, clsWriteProtectionItem)
        Get
            liste = _allWriteProtections
        End Get
        Set(value As SortedList(Of String, clsWriteProtectionItem))
            If Not IsNothing(value) Then
                _allWriteProtections = value
            Else
                _allWriteProtections = New SortedList(Of String, clsWriteProtectionItem)
            End If
        End Set
    End Property

    Public ReadOnly Property isProtected(ByVal pvname As String) As Boolean
        Get
            Dim tmpResult As Boolean = False
            If _allWriteProtections.ContainsKey(pvname) Then
                tmpResult = _allWriteProtections.Item(pvname).isProtected
            End If
            isProtected = tmpResult
        End Get
    End Property

    Public ReadOnly Property isPermanentProtected(ByVal pvName As String) As Boolean
        Get
            Dim tmpResult As Boolean = False
            If _allWriteProtections.ContainsKey(pvName) Then
                Dim tmpItem As clsWriteProtectionItem = _allWriteProtections.Item(pvName)
                tmpResult = tmpItem.isProtected And tmpItem.permanent
            End If
            isPermanentProtected = tmpResult
        End Get
    End Property

    Public ReadOnly Property wasProtectedBy(ByVal pvName As String) As String
        Get
            Dim tmpResult As String = ""
            If _allWriteProtections.ContainsKey(pvName) Then
                Dim tmpItem As clsWriteProtectionItem = _allWriteProtections.Item(pvName)
                If tmpItem.isProtected Then
                    tmpResult = _allWriteProtections.Item(pvName).userName
                End If
            End If
            wasProtectedBy = tmpResult
        End Get
    End Property

    Public ReadOnly Property wasReleasedBy(ByVal pvName As String) As String
        Get
            Dim tmpResult As String = ""
            If _allWriteProtections.ContainsKey(pvName) Then
                Dim tmpItem As clsWriteProtectionItem = _allWriteProtections.Item(pvName)
                If Not tmpItem.isProtected Then
                    tmpResult = _allWriteProtections.Item(pvName).userName
                End If
            End If
            wasReleasedBy = tmpResult
        End Get
    End Property

    Public Sub New()
        _allWriteProtections = New SortedList(Of String, clsWriteProtectionItem)
    End Sub
End Class
