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

    ''' <summary>
    ''' wenn ein userName angegeben ist, wird nur dann true zurückgegeben, wenn der userName ungleich dem User ist, der das Projekt geschützt hat 
    ''' wenn es der gleiche User ist, wird false zurückgegeben 
    ''' </summary>
    ''' <param name="pvname"></param>
    ''' <param name="userName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isProtected(ByVal pvname As String, Optional ByVal userName As String = Nothing) As Boolean
        Get
            Dim tmpResult As Boolean = False
            If _allWriteProtections.ContainsKey(pvname) Then
                If IsNothing(userName) Then
                    tmpResult = _allWriteProtections.Item(pvname).isProtected

                ElseIf userName <> "" Then
                    tmpResult = _allWriteProtections.Item(pvname).isProtected And _
                                _allWriteProtections.Item(pvname).userName <> userName

                Else
                    tmpResult = _allWriteProtections.Item(pvname).isProtected
                End If

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

    Public ReadOnly Property lastModifiedBy(ByVal pvName As String) As String
        Get
            Dim tmpResult As String = ""
            If _allWriteProtections.ContainsKey(pvName) Then
                Dim tmpItem As clsWriteProtectionItem = _allWriteProtections.Item(pvName)
                If tmpItem.isProtected Then
                    tmpResult = _allWriteProtections.Item(pvName).userName
                Else
                    ' auch beim Release wird ja der User-Name eingetragen 
                    tmpResult = _allWriteProtections.Item(pvName).userName
                End If
            End If
            lastModifiedBy = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' liefert das Datum zurück, wann das Item geschützt / released wurde
    ''' </summary>
    ''' <param name="pvName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property changeDate(ByVal pvName As String) As Date
        Get
            Dim tmpResult As Date
            If _allWriteProtections.ContainsKey(pvName) Then
                Dim tmpItem As clsWriteProtectionItem = _allWriteProtections.Item(pvName)

                If tmpItem.isProtected Then
                    tmpResult = _allWriteProtections.Item(pvName).lastDateSet
                Else
                    tmpResult = _allWriteProtections.Item(pvName).lastDateReleased
                End If

            End If
            changeDate = tmpResult
        End Get
    End Property

    Public Sub upsert(ByVal wpItem As clsWriteProtectionItem)

        If Not IsNothing(wpItem) Then
            If _allWriteProtections.ContainsKey(wpItem.pvName) Then
                ' update 
                _allWriteProtections.Item(wpItem.pvName) = wpItem
            Else
                ' insert 
                _allWriteProtections.Add(wpItem.pvName, wpItem)
            End If
        End If
        
    End Sub

    Public Sub New()
        _allWriteProtections = New SortedList(Of String, clsWriteProtectionItem)
    End Sub
End Class
