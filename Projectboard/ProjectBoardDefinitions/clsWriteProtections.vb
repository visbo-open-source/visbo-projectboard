Public Class clsWriteProtections
    Private _allWriteProtections As SortedList(Of String, clsWriteProtectionItem)

    Public ReadOnly Property getProtectionText(ByVal pvname As String) As String
        Get
            Dim tmpText As String = ""
            If Me.isProtected(pvname) Then
                Dim permanent As String = ""
                If Me.isPermanentProtected(pvname) Then
                    permanent = "permanent "
                End If
                If awinSettings.englishLanguage Then
                    tmpText = permanent & "protected by: " & Me.lastModifiedBy(pvname) & ", at: " & Me.changeDate(pvname).ToString
                Else
                    tmpText = permanent & "geschützt von: " & Me.lastModifiedBy(pvname) & ", am: " & Me.changeDate(pvname).ToString
                End If

            Else
                If awinSettings.englishLanguage Then
                    tmpText = "no protection"
                Else
                    tmpText = "nicht geschützt"
                End If
            End If
            getProtectionText = tmpText
        End Get
    End Property

    ''' <summary>
    ''' setzt die WriteProtections zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        _allWriteProtections.Clear()
    End Sub

    ''' <summary>
    ''' aktualisiert die bestehende Liste durch die neue Liste; alle Einträge, die nur Session-Projekte sind, bleiben unverändert 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property adjustListe As SortedList(Of String, clsWriteProtectionItem)

        Set(value As SortedList(Of String, clsWriteProtectionItem))
            If Not IsNothing(value) Then
                ' sicherstellen, dass die Projekt-Bezeichner in der richtigen Farbe / Font dargestellt werden 
                For Each kvp As KeyValuePair(Of String, clsWriteProtectionItem) In value
                    ' das aktualisiert jetzt ggf auch die Namen der Projekte auf der Multiprojekt-Tafel 
                    Call Me.upsert(kvp.Value)
                Next
                ' dieser Befehl darf nicht ausgeführt werden, weil sonst alle nur in der Session vorhandenen 
                ' Projekte in der Liste verloren gehen 
                ' jetzt einfach die komplette Liste umhängen ... 
                '_allWriteProtections = value
            Else
                ' Alle Einträge löschen bis auf die , die nur in der Session sind 
                For Each kvp As KeyValuePair(Of String, clsWriteProtectionItem) In _allWriteProtections
                    If kvp.Value.isSessionOnly Then
                        ' nichts tun 
                    Else
                        _allWriteProtections.Remove(kvp.Key)
                    End If
                Next
                ' tk, wurde durch das obige ersetzt ... 
                '_allWriteProtections = New SortedList(Of String, clsWriteProtectionItem)
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

                Else

                    tmpResult = _allWriteProtections.Item(pvname).isProtected And _
                                _allWriteProtections.Item(pvname).userName <> userName
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

    ''' <summary>
    ''' aktualisiert die Hauptspeicher-Struktur der Schreibberechtigungen und aktualisiert das Erscheinungsbild auf der Multiprojekt-Tafel 
    ''' </summary>
    ''' <param name="wpItem"></param>
    ''' <remarks></remarks>
    Public Sub upsert(ByVal wpItem As clsWriteProtectionItem)

        If Not IsNothing(wpItem) Then
            ' prüfen, ob sich der Protect status ändert , wenn ja, soll auch gleich die Projekt-Namen Änderung angestossen werden 

            If _allWriteProtections.ContainsKey(wpItem.pvName) Then

                Dim chkItem As clsWriteProtectionItem = _allWriteProtections.Item(wpItem.pvName)
                ' jetzt updaten 
                _allWriteProtections.Item(wpItem.pvName) = wpItem

                ' muss die Darstellung auf der Multiprojekt-Tafel upgedated werden ? 
                If chkItem.isProtected <> wpItem.isProtected Or _
                    ((chkItem.isProtected = wpItem.isProtected) And (chkItem.permanent <> wpItem.permanent)) Or
                    ((chkItem.isProtected = wpItem.isProtected) And (chkItem.userName <> wpItem.userName)) Then
                    ' auf der Multiprojekt-Tafel muss der Name aktualisiert werden  
                    Dim pName As String = getPnameFromKey(wpItem.pvName)
                    Dim vName As String = getVariantnameFromKey(wpItem.pvName)
                    If ShowProjekte.contains(pName) Then
                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Call zeichneNameInProjekt(hproj)
                    End If
                End If

            Else
                ' insert 
                _allWriteProtections.Add(wpItem.pvName, wpItem)

                ' muss die Darstellung auf der Multiprojekt-Tafel upgedated werden ? 
                Dim pName As String = getPnameFromKey(wpItem.pvName)
                Dim vName As String = getVariantnameFromKey(wpItem.pvName)
                If ShowProjekte.contains(pName) Then
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    Call zeichneNameInProjekt(hproj)
                End If
            End If
        End If

    End Sub

    Public Sub New()
        _allWriteProtections = New SortedList(Of String, clsWriteProtectionItem)
    End Sub
End Class
