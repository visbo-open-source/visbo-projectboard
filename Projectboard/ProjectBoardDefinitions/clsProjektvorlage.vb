Public Class clsProjektvorlage

   

    Public AllPhases As List(Of clsPhase)

    ' Änderung tk 31.3.15 Hierachie Klasse ergänzt 
    Public hierarchy As clsHierarchy

    ' Änderung tk 20.9.16 sortierte Listen, wo welche Rollen vorkommen ... 
    Public rcLists As clsListOfCostAndRoles



    Private relStart As Integer
    Private uuid As Long
    ' als Friend deklariert, damit sie aus der Klasse clsProjekt, die von clsProjektvorlage erbt , erreichbar ist
    Friend _Dauer As Integer
    Private _earliestStart As Integer
    Private _latestStart As Integer

    ' Hinzufügen von Custom Feldern beliebiger Anzahl 
    ' ein CustomFeld eines bestimmten Typs darf nur einmal vorkommen 
    Private _customDblFields As SortedList(Of Integer, Double)
    Private _customStringFields As SortedList(Of Integer, String)
    Private _customBoolFields As SortedList(Of Integer, Boolean)


    ''' <summary>
    ''' gibt die sortierte Liste der Double Customfields zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property customDblFields As SortedList(Of Integer, Double)
        Get

            If IsNothing(_customDblFields) Then
                _customDblFields = New SortedList(Of Integer, Double)
            End If

            customDblFields = _customDblFields

        End Get
    End Property

    ''' <summary>
    ''' gibt die sortierte Liste der String Customfields zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property customStringFields As SortedList(Of Integer, String)
        Get

            If IsNothing(_customStringFields) Then
                _customStringFields = New SortedList(Of Integer, String)
            End If

            customStringFields = _customStringFields

        End Get
    End Property

    ''' <summary>
    ''' gibt die sortierte Liste der bool'schen Customfields zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property customBoolFields As SortedList(Of Integer, Boolean)
        Get

            If IsNothing(_customBoolFields) Then
                _customBoolFields = New SortedList(Of Integer, Boolean)
            End If

            customBoolFields = _customBoolFields

        End Get
    End Property

    ''' <summary>
    ''' gibt den Wert für das Double Custom-Field mit Identifier UID zurück; wenn das Custom Field nicht existiert, wird Nothing zurückgegeben
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomDField(ByVal uid As Integer) As Double
        Get
            Dim tmpValue As Double = Nothing

            If _customDblFields.ContainsKey(uid) Then
                tmpValue = _customDblFields.Item(uid)
            End If

            getCustomDField = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' gibt den Wert des Double Custom-Fields mit NAme cfName zurück; Nothing, wenn es nicht existiert 
    ''' </summary>
    ''' <param name="cfName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomDField(ByVal cfName As String) As Double
        Get
            Dim tmpValue As Double = Nothing
            Dim uid As Integer = customFieldDefinitions.getUid(cfName, ptCustomFields.Dbl)

            If _customDblFields.ContainsKey(uid) Then
                tmpValue = _customDblFields.Item(uid)
            End If

            getCustomDField = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' fügt der Liste an Double CustomFields ein neues hinzu
    ''' wenn das CustomField gar nicht definiert ist: Exception
    ''' wenn das Feld schon existiert, dann wird der Wert aktualisiert
    ''' wenn Nothing als Wert übergeben wird, wird der Default 0.0  angenommen 
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub addSetCustomDField(ByVal uid As Integer, ByVal value As Double)


        ' wenn etwas schief geht, bleibt es auf false
        Dim ok As Boolean = False

        If IsNothing(value) Then
            value = 0.0
        End If

        ' ist es überhaupt eine gültige uid?  
        If customFieldDefinitions.contains(uid) Then
            ' wenn es das Custom field in dem Projekt schon gibt 
            If customFieldDefinitions.getDef(uid).type = ptCustomFields.Dbl Then
                ' nur dann soll das gesetzt werden ...
                ok = True
                If _customDblFields.ContainsKey(uid) Then
                    _customDblFields.Item(uid) = value
                Else
                    _customDblFields.Add(uid, value)
                End If
            End If
        End If

        If Not ok Then
            Throw New ArgumentException("uid nicht bekannt oder hat falschen Typ (nicht Dbl):" & uid.ToString)
        End If


    End Sub


    ''' <summary>
    ''' gibt den Wert für das String Custom-Field mit Identifier UID zurück; wenn das Custom Field nicht existiert, wird Nothing zurückgegeben
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomSField(ByVal uid As Integer) As String
        Get
            Dim tmpValue As String = Nothing

            If _customStringFields.ContainsKey(uid) Then
                tmpValue = _customStringFields.Item(uid)
            End If

            getCustomSField = tmpValue

        End Get
    End Property


    
    ''' <summary>
    ''' gibt den Wert des String Custom-Fields mit Name cfName zurück; Nothing, wenn es nicht existiert
    ''' </summary>
    ''' <param name="cfName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomSField(ByVal cfName As String) As String
        Get
            Dim tmpValue As String = Nothing
            Dim uid As Integer = customFieldDefinitions.getUid(cfName, ptCustomFields.Str)

            If _customStringFields.ContainsKey(UID) Then
                tmpValue = _customStringFields.Item(UID)
            End If

            getCustomSField = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' fügt der Liste an String CustomFields ein neues hinzu
    ''' wenn das Feld schon existiert, dann wird der Wert aktualisiert
    ''' wenn Nothing als Wert übergeben wird, wird der Default "?"  angenommen 
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub addSetCustomSField(ByVal uid As Integer, ByVal value As String)

        ' wenn etwas schief geht, bleibt es auf false
        Dim ok As Boolean = False

        If IsNothing(value) Then
            value = ""
        End If

        ' ist es überhaupt eine gültige uid?  
        If customFieldDefinitions.contains(uid) Then
            ' wenn es das Custom field in dem Projekt schon gibt 
            If customFieldDefinitions.getDef(uid).type = ptCustomFields.Str Then
                ' nur dann soll das gesetzt werden ...
                ok = True
                If _customStringFields.ContainsKey(uid) Then
                    _customStringFields.Item(uid) = value
                Else
                    _customStringFields.Add(uid, value)
                End If
            End If
        End If

        If Not ok Then
            Throw New ArgumentException("uid nicht bekannt oder hat falschen Typ (nicht String):" & uid.ToString)
        End If

    End Sub

    ''' <summary>
    ''' gibt den Wert für das Bool Custom-Field mit  Namen key zurück; wenn das Custom Field nicht existiert, wird Nothing zurückgegeben
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomBField(ByVal uid As Integer) As Boolean
        Get
            Dim tmpValue As Boolean = Nothing

            If _customBoolFields.ContainsKey(uid) Then
                tmpValue = _customBoolFields.Item(uid)
            End If

            getCustomBField = tmpValue

        End Get
    End Property


    ''' <summary>
    ''' gibt den Wert des Bool Custom-Fields mit NAme cfName zurück; Nothing, wenn es nicht existiert 
    ''' </summary>
    ''' <param name="cfName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomBField(ByVal cfName As String) As Boolean
        Get
            Dim tmpValue As Boolean = Nothing
            Dim uid As Integer = customFieldDefinitions.getUid(cfName, ptCustomFields.bool)

            If _customBoolFields.ContainsKey(uid) Then
                tmpValue = _customBoolFields.Item(uid)
            End If

            getCustomBField = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' fügt der Liste an bool'schen CustomFields ein neues hinzu
    ''' wenn das CustomField gar nicht definiert ist: Exception
    ''' wenn das Feld schon existiert, dann wird der Wert aktualisiert
    ''' wenn Nothing als Wert übergeben wird, wird der Default Wert false angenommen 
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub addSetCustomBField(ByVal uid As Integer, ByVal value As Boolean)

        ' wenn etwas schief geht, bleibt es auf false
        Dim ok As Boolean = False

        If IsNothing(value) Then
            value = False
        End If

        ' ist es überhaupt eine gültige uid?  
        If customFieldDefinitions.contains(uid) Then
            ' wenn es das Custom field in dem Projekt schon gibt 
            If customFieldDefinitions.getDef(uid).type = ptCustomFields.bool Then
                ' nur dann soll das gesetzt werden ...
                ok = True
                If _customBoolFields.ContainsKey(uid) Then
                    _customBoolFields.Item(uid) = value
                Else
                    _customBoolFields.Add(uid, value)
                End If
            End If
        End If

        If Not ok Then
            Throw New ArgumentException("uid nicht bekannt oder hat falschen Typ (nicht bool):" & uid.ToString)
        End If


    End Sub

    ''' <summary>
    ''' gibt den kürzesten eindeutigen Namen für das Element zurück, der sich finden lässt
    ''' optional kann die SwimlaneID mitgegeben werden - dann wird nur nach eindeutigen Namen innerhalb der swimlanes gesucht 
    ''' wenn das Element eh eindeutig ist im Projekt, dann wird nur der Elem-Name zurückgegeben 
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBestNameOfID(ByVal nameID As String, _
                                             ByVal ShowStdNames As Boolean, ByVal showAbbrev As Boolean, _
                                             Optional ByVal swimlaneID As String = rootPhaseName) As String
        Get
            Dim elemName As String = elemNameOfElemID(nameID)
            Dim isMilestone As Boolean
            Dim curBC As String = ""
            Dim oldBC As String = ""
            Dim anzElements As Integer
            Dim anzElementsBefore As Integer
            Dim level As Integer = 1
            Dim tmpName As String = elemName
            Dim rootreached As Boolean = False
            Dim description1 As String = "", description2 As String = elemName
            Dim phDef As clsPhasenDefinition
            Dim swlBC As String = ""



            isMilestone = elemIDIstMeilenstein(nameID)

            If swimlaneID = rootPhaseName Then
                swlBC = ""
            Else
                If istElemID(swimlaneID) Then
                    swlBC = calcHryFullname(elemNameOfElemID(swimlaneID), _
                                                  Me.hierarchy.getBreadCrumb(swimlaneID))
                End If
            End If

            Try
                If isMilestone Then

                    ' Änderung tk: es wird der eindeutige Namen unterhalb der swimlaneID gesucht  
                    'Dim milestoneIndices(,) As Integer = Me.getMilestoneIndices(elemName, "")
                    Dim milestoneIndices(,) As Integer = Me.hierarchy.getMilestoneIndices(elemName, swlBC)
                    anzElements = CInt(milestoneIndices.Length / 2)

                    If anzElements > 1 Then

                        anzElementsBefore = anzElements

                        Do Until anzElements = 1 Or rootreached
                            curBC = Me.hierarchy.getBreadCrumb(nameID, level)

                            If oldBC = curBC Then
                                rootreached = True
                            Else
                                oldBC = curBC
                            End If

                            If Not rootreached Then
                                milestoneIndices = Me.hierarchy.getMilestoneIndices(elemName, curBC)
                                anzElements = CInt(milestoneIndices.Length / 2)
                                If anzElements < anzElementsBefore Then
                                    anzElementsBefore = anzElements
                                    tmpName = calcHryFullname(elemName, curBC)
                                End If
                            End If

                            level = level + 1

                        Loop
                    Else
                        tmpName = elemName
                    End If


                Else
                    ' es handelt sich um eine Phase
                    'Dim phaseIndices() As Integer = Me.getPhaseIndices(elemName, "")
                    ' Änderung tk: es wird der eindeutige Namen unterhalb der swimlaneID gesucht  
                    Dim phaseIndices() As Integer = Me.hierarchy.getPhaseIndices(elemName, swlBC)
                    anzElements = phaseIndices.Length

                    If anzElements > 1 Then

                        anzElementsBefore = anzElements

                        Do Until anzElements = 1 Or rootreached
                            curBC = Me.hierarchy.getBreadCrumb(nameID, level)

                            If oldBC = curBC Then
                                rootreached = True
                            Else
                                oldBC = curBC
                            End If

                            If Not rootreached Then
                                phaseIndices = Me.hierarchy.getPhaseIndices(elemName, curBC)
                                anzElements = phaseIndices.Length
                                If anzElements < anzElementsBefore Then
                                    anzElementsBefore = anzElements
                                    tmpName = calcHryFullname(elemName, curBC)
                                End If
                            End If

                            level = level + 1

                        Loop
                    Else
                        tmpName = elemName
                    End If
                End If
            Catch ex As Exception

            End Try

            ' jetzt wird unterschieden, ob Abbrev gezeigt werden soll oder Standard Name ... 
            If ShowStdNames Then
                If showAbbrev Then

                    If awinSettings.showBestName And Not awinSettings.drawphases Then
                        ' den bestmöglichen, also den kürzesten Breadcrumb Namen, der (möglichst) eindeutig ist
                        ' anzeigen; aber nur, wenn im Ein-Zeile-Modus beschriftet wird, weil dann der Kontext fehlt ... 
                        Dim type As Integer = -1
                        Dim pvName As String = ""
                        Call splitHryFullnameTo2(tmpName, description2, description1, type, pvName)

                        Dim tmpStr() As String = description1.Split(New Char() {CChar("#")}, 20)

                        ' jetzt den Abbrev String zusammensetzen 
                        Dim newDesc1 As String = ""
                        For i As Integer = 1 To tmpStr.Length
                            Dim tmpPhName As String = tmpStr(i - 1)
                            phDef = PhaseDefinitions.getPhaseDef(tmpPhName)

                            If IsNothing(phDef) Then
                                If i = 1 And tmpPhName <> elemNameOfElemID(rootPhaseName) And tmpPhName <> "" Then
                                    ' den tmpPhName eintragen 
                                    newDesc1 = tmpPhName
                                ElseIf i > 1 Then
                                    newDesc1 = newDesc1 & tmpPhName
                                End If

                            Else
                                If i = 1 Then
                                    If phDef.shortName = "" Then
                                        newDesc1 = tmpPhName
                                    Else
                                        newDesc1 = phDef.shortName
                                    End If

                                Else
                                    If phDef.shortName = "" Then
                                        newDesc1 = newDesc1 & tmpPhName
                                    Else
                                        newDesc1 = newDesc1 & "-" & phDef.shortName
                                    End If

                                End If
                            End If

                        Next
                        description1 = newDesc1

                        If isMilestone Then

                            Dim msDef As clsMeilensteinDefinition
                            msDef = MilestoneDefinitions.getMilestoneDef(description2)
                            If IsNothing(msDef) Then
                                msDef = missingMilestoneDefinitions.getMilestoneDef(description2)
                            End If

                            If IsNothing(msDef) Then
                                ' nichts zu tun
                            Else
                                If IsNothing(msDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If msDef.shortName = "" Then
                                        'description2 = "-"
                                    Else
                                        description2 = msDef.shortName
                                    End If

                                End If
                            End If

                        Else

                            phDef = PhaseDefinitions.getPhaseDef(description2)
                            If IsNothing(phDef) Then

                                phDef = missingPhaseDefinitions.getPhaseDef(description2)
                            End If

                            If IsNothing(phDef) Then
                                ' nichts zu tun
                            Else

                                If IsNothing(phDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If phDef.shortName = "" Then
                                        'description2 = "-"
                                    Else
                                        description2 = phDef.shortName
                                    End If

                                End If
                            End If

                        End If
                    Else
                        description1 = ""

                        If isMilestone Then

                            Dim msDef As clsMeilensteinDefinition
                            msDef = MilestoneDefinitions.getMilestoneDef(description2)
                            If IsNothing(msDef) Then
                                msDef = missingMilestoneDefinitions.getMilestoneDef(description2)
                            End If

                            If IsNothing(msDef) Then
                                'description2 = "-"
                            Else

                                If IsNothing(msDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If msDef.shortName = "" Then
                                        'description2 = msDef.name
                                    Else
                                        description2 = msDef.shortName
                                    End If

                                End If
                            End If

                        Else

                            phDef = PhaseDefinitions.getPhaseDef(description2)
                            If IsNothing(phDef) Then

                                phDef = missingPhaseDefinitions.getPhaseDef(description2)
                            End If

                            If IsNothing(phDef) Then
                                'description2 = "-"
                            Else

                                If IsNothing(phDef.shortName) Then
                                    'description2 = "-"
                                Else
                                    If phDef.shortName = "" Then
                                        'description2 = phDef.name
                                    Else
                                        description2 = phDef.shortName
                                    End If

                                End If
                            End If

                        End If


                    End If
                Else
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitHryFullnameTo2(tmpName, description2, description1, type, pvName)
                    Dim tmpStr() As String = description1.Split(New Char() {CChar("#")}, 20)

                    ' jetzt den Std-Name zusammensetzen 
                    Dim newDesc1 As String = ""
                    For i As Integer = 1 To tmpStr.Length
                        Dim tmpPhName As String = tmpStr(i - 1)

                        If i = 1 Then
                            If tmpPhName = elemNameOfElemID(rootPhaseName) Then
                                ' nichts tun
                            Else
                                newDesc1 = tmpPhName
                            End If
                        ElseIf i > 1 Then
                            newDesc1 = newDesc1 & "-" & tmpPhName
                        End If

                    Next
                    description1 = newDesc1


                End If
            Else
                If isMilestone Then
                    description2 = Me.getMilestoneByID(nameID).originalName
                Else
                    description2 = Me.getPhaseByID(nameID).originalName
                End If
            End If

            Dim description As String = ""
            Try
                If description1 <> "" Then
                    description = description1 & "-" & description2
                Else
                    description = description2
                End If
            Catch ex As Exception

            End Try

            getBestNameOfID = description

        End Get
    End Property


    ''' <summary>
    ''' Bezugsdatum ist hier der StartofCalendar
    ''' während in der addphase der abgeleiteten ProjektKlasse das Projektstartdatum das maßgebliche Datum ist 
    ''' </summary>
    ''' <param name="phase"></param>
    ''' <remarks></remarks>
    Public Overridable Sub AddPhase(ByVal phase As clsPhase, _
                                    Optional ByVal origName As String = "", _
                                    Optional ByVal parentID As String = "")

        Dim phaseEnde As Double
        Dim maxM As Integer


        ' wenn der Origname gesetzt werden soll ...
        If origName <> "" Then
            If phase.originalName <> origName Then
                phase.originalName = origName
            End If
        End If

        With phase

            phaseEnde = .startOffsetinDays + .dauerInDays - 1

        End With

        If phaseEnde > 0 Then

            maxM = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(phaseEnde)) + 1)
            If maxM <> _Dauer And maxM > 0 Then
                _Dauer = maxM
                ' hier muss jetzt die Dauer der Allgemeinen Phase angepasst werden ... 
            End If
        End If


        AllPhases.Add(phase)

        ' jetzt muss die Phase in die Projekt-Hierarchie aufgenommen werden 
        Dim currentElementNode As New clsHierarchyNode
        With currentElementNode

            If Me.CountPhases = 1 Then
                .elemName = "."
            Else
                .elemName = phase.name
            End If

            ' Änderung tk 29.5.16 origName ist nicht mehr Bestandteil von HierarchyNode, 
            ' sondern von clsPhase
            'If origName = "" Then
            '    .origName = .elemName
            'Else
            '    .origName = origName
            'End If

            .indexOfElem = Me.CountPhases

            If parentID = "" Then
                If .indexOfElem = 1 Then
                    .parentNodeKey = ""
                Else
                    .parentNodeKey = calcHryElemKey(".", False)
                End If
            Else
                .parentNodeKey = parentID
            End If

        End With

        With Me.hierarchy
            .addNode(currentElementNode, phase.nameID)
        End With


    End Sub

    ''' <summary>
    ''' entfernt die Phase mit der übergebenen nameID 
    ''' dabei kann angegeben werden, was mit den Kind-Elementen passieren soll: löschen oder umhängen 
    ''' die rootPhase kann nicht gelöscht werden; in diesem Fall wird eine Exception geworfen  
    ''' </summary>
    ''' <param name="nameID">der eindeutige Identifier aus der Hierarchie-Liste</param>
    ''' <param name="deleteAllChilds" >
    ''' true: alle Kind-Elemente werden mitgelöscht
    ''' false: alle Kind-Elemente werden der Parent-Phase zugewiesen  </param>
    ''' <remarks></remarks>
    Public Sub removePhase(ByVal nameID As String, Optional deleteAllChilds As Boolean = True)

        ' die Root-Phase darf nicht gelöscht werden ...
        If nameID = rootPhaseName Then
            Throw New ArgumentException(message:="die Root-Phase kann nicht gelöscht werden  ", paramName:=nameID)
        End If

        If elemIDIstMeilenstein(nameID) Then
            Throw New ArgumentException(message:="das übergebene Element ist keine Phase ... ", paramName:=nameID)
        End If

        Dim elemNode As clsHierarchyNode = Me.hierarchy.nodeItem(nameID)

        ' Abbruch, wenn das Element gar nicht existiert 
        If IsNothing(elemNode) Then
            Throw New ArgumentException(message:="das Element existiert nicht in der Hierarchie: ", paramName:=nameID)
        End If

        ' Konsistenzprüfung: stimmt der Verweis ? 
        Dim indexInPhaseList As Integer = elemNode.indexOfElem
        If Me.AllPhases.ElementAt(indexInPhaseList - 1).nameID <> nameID Then
            Throw New ArgumentException(message:="der Verweis auf die Phasen-Liste ist nicht korrekt ", paramName:=nameID)
        End If


        Dim parentID As String = elemNode.parentNodeKey
        Dim parentNode As clsHierarchyNode = Me.hierarchy.parentNodeItem(nameID)
        Dim childNodeID As String = ""

        'als erstes im ParentNode das Element aus der Kinder-Liste löschen 
        parentNode.removeChild(nameID)

        If deleteAllChilds Then

            Dim k As Integer = 1


            ' jetzt alle Kinder löschen
            While elemNode.childCount > 0

                childNodeID = elemNode.getChild(k)
                If elemIDIstMeilenstein(childNodeID) Then
                    ' lösche Meilenstein 
                    Me.removeMeilenstein(childNodeID)
                Else
                    Me.removePhase(childNodeID, True)

                End If
            End While

            '' '' ''For i As Integer = 1 To elemNode.childCount
            '' '' ''    childNodeID = elemNode.getChild(i)
            '' '' ''    If elemIDIstMeilenstein(childNodeID) Then
            '' '' ''         lösche Meilenstein 
            '' '' ''        Me.removeMeilenstein(childNodeID)
            '' '' ''    Else
            '' '' ''        Me.removePhase(childNodeID, True)
            '' '' ''    End If
            '' '' ''Next
        Else
            ' hier alle Kinder umhängen: die bekommen die ParentID statt nameID als ihren neuen Vater 
            For i As Integer = 1 To elemNode.childCount
                Dim childNode As clsHierarchyNode
                childNodeID = elemNode.getChild(i)
                If Me.hierarchy.containsKey(childNodeID) Then
                    childNode = Me.hierarchy.nodeItem(childNodeID)
                    childNode.parentNodeKey = parentID
                End If
            Next
        End If

        Dim indexInHierarchy As Integer = Me.hierarchy.getIndexOfID(nameID)

        ' in der Hierarchie-Liste löschen 
        Me.hierarchy.removeAt(indexInHierarchy - 1)

        ' in der Phasen-Liste löschen
        Me.AllPhases.RemoveAt(indexInPhaseList - 1)

        ' jetzt in der Hierarchie alle Phasen-Verweise, die größer als indexInPhaseList sind, um eins erniedrigen 
        Me.hierarchy.updatePhasenVerweise(indexInPhaseList, -1)


    End Sub

   

    ''' <summary>
    ''' entfernt den Meilenstein mit der übergebenen nameID 
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <remarks></remarks>
    Public Sub removeMeilenstein(ByVal nameID As String)

        If Not elemIDIstMeilenstein(nameID) Then
            Throw New ArgumentException(message:="das übergebene Element ist kein Meilenstein ... ", paramName:=nameID)
        End If


        Dim elemNode As clsHierarchyNode = Me.hierarchy.nodeItem(nameID)

        ' Abbruch, wenn das Element gar nicht existiert 
        If IsNothing(elemNode) Then
            Throw New ArgumentException(message:="das Element existiert nicht in der Hierarchie: ", paramName:=nameID)
        End If

        Dim parentID As String = elemNode.parentNodeKey
        Dim parentNode As clsHierarchyNode = Me.hierarchy.parentNodeItem(nameID)
        Dim childNodeID As String = ""

        'als erstes im ParentNode das Element aus der Kinder-Liste löschen 
        parentNode.removeChild(nameID)

        ' ein Meilenstein kann eigentlich keine Kinder haben, Fehler, wenn doch ..
        If elemNode.childCount > 0 Then
            Call MsgBox("Meilenstein mit Kindern !?")
        End If

        ' jetzt den Meilenstein selber löschen 
        Dim indexInMilestoneList As Integer = elemNode.indexOfElem
        Dim indexInHierarchy As Integer = Me.hierarchy.getIndexOfID(nameID)

        ' in der Hierarchie-Liste löschen 
        Me.hierarchy.removeAt(indexInHierarchy - 1)

        Dim cPhase As clsPhase = Me.getPhaseByID(parentID)

        ' in der Meilenstein-Liste der Phase löschen 
        cPhase.removeMilestoneAt(indexInMilestoneList - 1)

        ' jetzt in der Hierarchie alle Meilenstein-Verweise, die größer als indexInMilestoneList sind, um eins erniedrigen 
        Me.hierarchy.updateMeilensteinVerweise(indexInMilestoneList, parentID, -1)


    End Sub

    ''' <summary>
    ''' gibt den Meilenstein mit Element-ID elemID zurück 
    ''' Nothing, wenn der Meilenstein nicht existiert 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneByID(elemID As String) As clsMeilenstein
        Get
            Dim currentNode As clsHierarchyNode = Me.hierarchy.nodeItem(elemID)
            Dim phaseID As String
            Dim phIndex As Integer, msIndex As Integer

            If Not IsNothing(currentNode) Then

                If elemIDIstMeilenstein(elemID) Then
                    phaseID = currentNode.parentNodeKey
                    phIndex = Me.hierarchy.nodeItem(phaseID).indexOfElem
                    msIndex = currentNode.indexOfElem

                    Dim cphase As clsPhase = Me.getPhase(phIndex)
                    If Not IsNothing(cphase) Then
                        getMilestoneByID = cphase.getMilestone(msIndex)
                    Else
                        getMilestoneByID = Nothing
                    End If
                Else
                    getMilestoneByID = Nothing
                End If

            Else
                getMilestoneByID = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die Parent-Phase zu der angegebenen Elem-ID zurück; 
    ''' wenn es keine Parent-Phase gibt oder 
    ''' wenn es das Element gar nicht gibt, wird Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getParentPhaseByID(ByVal elemID As String) As clsPhase
        Get

            Dim currentNode As clsHierarchyNode
            Dim phaseID As String
            Dim phIndex As Integer


            currentNode = Me.hierarchy.nodeItem(elemID)

            If Not IsNothing(currentNode) Then

                phaseID = currentNode.parentNodeKey
                phIndex = Me.hierarchy.nodeItem(phaseID).indexOfElem
                getParentPhaseByID = Me.getPhase(phIndex)

            Else
                getParentPhaseByID = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt zu einem gegebenen Meilenstein-Namen das clsMeilenstein Objekt zurück, sofern es existiert
    ''' Nothing sonst
    ''' </summary>
    ''' <param name="msName">Name des Meilensteins</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestone(ByVal msName As String, Optional ByVal breadcrumb As String = "", Optional ByVal lfdNr As Integer = 1) As clsMeilenstein
        Get
            Dim tmpMilestone As clsMeilenstein = Nothing
            Dim found As Boolean = False

            Dim milestoneIndices(,) As Integer
            milestoneIndices = Me.hierarchy.getMilestoneIndices(msName, breadcrumb)

            If lfdNr > CInt(milestoneIndices.Length / 2) Or lfdNr < 1 Then
                ' kein gültiger Meilenstein 
            Else
                If milestoneIndices(0, lfdNr - 1) > 0 And milestoneIndices(1, lfdNr - 1) > 0 Then
                    ' nur dann existiert dieser Meilenstein
                    tmpMilestone = Me.getMilestone(milestoneIndices(0, lfdNr - 1), milestoneIndices(1, lfdNr - 1))
                End If

            End If

            getMilestone = tmpMilestone


        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="msName"></param>
    ''' <param name="breadcrumb"></param>
    ''' <param name="lfdNr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneOffsetToProjectStart(ByVal msName As String, Optional ByVal breadcrumb As String = "", Optional ByVal lfdNr As Integer = 1) As Long
        Get
            Dim tmpMilestone As clsMeilenstein = Nothing
            Dim found As Boolean = False

            Dim milestoneIndices(,) As Integer
            milestoneIndices = Me.hierarchy.getMilestoneIndices(msName, breadcrumb)

            If lfdNr > CInt(milestoneIndices.Length / 2) Or lfdNr < 1 Then
                ' kein gültiger Meilenstein 
            Else
                If milestoneIndices(0, lfdNr - 1) > 0 And milestoneIndices(1, lfdNr - 1) > 0 Then
                    ' nur dann existiert dieser Meilenstein
                    tmpMilestone = Me.getMilestone(milestoneIndices(0, lfdNr - 1), milestoneIndices(1, lfdNr - 1))
                End If

            End If

            getMilestoneOffsetToProjectStart = tmpMilestone.offset + tmpMilestone.Parent.startOffsetinDays


        End Get
    End Property

    ''' <summary>
    ''' gibt den Meilenstein zurück, der in der Phase mit Index PhaseIndex, 
    ''' und dort in der Meilenstein Liste mit Index milestoneIndex vorkommt  
    ''' </summary>
    ''' <param name="phaseIndex"></param>
    ''' <param name="milestoneIndex"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestone(ByVal phaseIndex As Integer, ByVal milestoneIndex As Integer) As clsMeilenstein
        Get

            Dim tmpResult As clsMeilenstein = Nothing
            Dim found As Boolean = False

            If phaseIndex >= 1 And phaseIndex <= AllPhases.Count Then
                Dim cphase As clsPhase = AllPhases.Item(phaseIndex - 1)
                If milestoneIndex >= 1 And milestoneIndex <= cphase.countMilestones Then
                    tmpResult = cphase.getMilestone(milestoneIndex)
                End If
            End If


            getMilestone = tmpResult


        End Get
    End Property


    Private _farbe As Integer = RGB(220, 220, 220)
    Public Property farbe() As Integer
        Get
            If Not IsNothing(_farbe) Then
                farbe = _farbe
            Else
                farbe = 10
            End If
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                Try
                    _farbe = CInt(value)
                Catch ex As Exception
                    _farbe = 10
                End Try

            Else
                _farbe = 10
            End If

        End Set
    End Property

    Private _Schrift As Integer = 10
    Public Property Schrift() As Integer
        Get
            If Not IsNothing(_Schrift) Then
                Schrift = _Schrift
            Else
                Schrift = 10
            End If
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                Try
                    _Schrift = CInt(value)
                    If _Schrift >= 5 And _Schrift <= 20 Then
                        ' ok , nichts tun
                    Else
                        _Schrift = 10
                    End If
                Catch ex As Exception
                    _Schrift = 10
                End Try

            Else
                _Schrift = 10
            End If

        End Set
    End Property


    Private _Schriftfarbe As Integer = RGB(5, 5, 5)
    Public Property Schriftfarbe() As Object
        Get
            If Not IsNothing(_Schriftfarbe) Then
                Schriftfarbe = _Schriftfarbe
            Else
                Schriftfarbe = RGB(5, 5, 5)
            End If
        End Get
        Set(value As Object)
            If Not IsNothing(value) Then
                Try
                    _Schriftfarbe = CInt(value)
                Catch ex As Exception
                    _Schriftfarbe = RGB(5, 5, 5)
                End Try

            Else
                _Schriftfarbe = RGB(5, 5, 5)
            End If

        End Set
    End Property

    Private _VorlagenName As String = ""
    Public Property VorlagenName() As String
        Get
            If Not IsNothing(_VorlagenName) Then
                VorlagenName = _VorlagenName
            Else
                VorlagenName = ""
            End If
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _VorlagenName = value
            Else
                _VorlagenName = ""
            End If

        End Set
    End Property

    'Public RessourcenDefinitionsBereich As String

    'Public KostenDefinitionsBereich As String

    ''' <summary>
    ''' kopiert die Attribute einer Projektvorlage in newproject;  bei der Quelle handelt es sich um eine 
    ''' Vorlage  
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <remarks></remarks>
    ''' 
    Public Overridable Sub copyAttrTo(ByRef newproject As clsProjekt)

        With newproject
            .farbe = Me.farbe
            .Schrift = Me.Schrift
            .Schriftfarbe = Me.Schriftfarbe
            .VorlagenName = Me.VorlagenName
            .earliestStart = _earliestStart
            .latestStart = _latestStart
            .name = ""
        End With

        ' jetzt wird die Hierarchie kopiert 
        Call copyHryTo(newproject)

        ' jetzt werden die CustomFields kopiert, so fern es welche gibt ... 
        Try
            With newproject
                For Each kvp As KeyValuePair(Of Integer, String) In Me.customStringFields
                    .customStringFields.Add(kvp.Key, kvp.Value)
                Next

                For Each kvp As KeyValuePair(Of Integer, Double) In Me.customDblFields
                    .customDblFields.Add(kvp.Key, kvp.Value)
                Next

                For Each kvp As KeyValuePair(Of Integer, Boolean) In Me.customBoolFields
                    .customBoolFields.Add(kvp.Key, kvp.Value)
                Next

            End With
        Catch ex As Exception

        End Try
        


    End Sub

    Public Overridable Sub copyTo(ByRef newproject As clsProjekt)
        Dim p As Integer
        Dim newphase As clsPhase
        Dim oldPhase As clsPhase
        'Dim parentID As String
        Dim origName As String = ""

        Call copyAttrTo(newproject)

        For p = 0 To Me.CountPhases - 1
            oldPhase = AllPhases.Item(p)
            newphase = New clsPhase(newproject)

            'parentID = Me.hierarchy.getParentIDOfID(oldPhase.nameID)

            oldPhase.copyTo(newphase)
            newproject.AddPhase(newphase)
            'newproject.AddPhase(newphase, origName:="", parentID:=parentID)
        Next p


    End Sub


    Public Overridable Sub korrCopyTo(ByRef newproject As clsProjekt, ByVal startdate As Date, ByVal endedate As Date, _
                                      Optional ByVal zielRenditenVorgabe As Double = -99999.0)
        Dim p As Integer
        Dim newphase As clsPhase
        Dim ProjectDauerInDays As Integer
        Dim CorrectFactor As Double
        Dim newPhaseNameID As String = ""

        Call copyAttrTo(newproject)
        newproject.startDate = startdate


        ProjectDauerInDays = calcDauerIndays(startdate, endedate)
        CorrectFactor = ProjectDauerInDays / Me.dauerInDays

        For p = 0 To Me.CountPhases - 1
            newphase = New clsPhase(newproject)
            AllPhases.Item(p).korrCopyTo(newphase, CorrectFactor, newPhaseNameID, zielRenditenVorgabe)

            newproject.AddPhase(newphase)
        Next p


    End Sub

    ''' <summary>
    ''' kopiert ein existierendes Modul; 
    ''' wenn moduleName ungleich "" dann wird noch eine Phase mit Dauer moduleDauerinDays angelegt  
    ''' </summary>
    ''' <param name="project">gibt das Projekt an, unter dem das Modul angelegt werden soll</param>
    ''' <param name="parentID">gibt die Parent-ID an, unter der das Modul angelegt werden soll</param>
    ''' <param name="moduleName">wenn ein Name angegeben ist, wird eine übergeordnete Phase mit diesem Namen angelegt </param>
    ''' <param name="modulStartoffset">gibt die Anzahl Tage an, die der Start des Moduls vom ProjektStart entfernt ist</param>
    ''' <param name="endoffset">gibt die Anzahl Tage an, die das Ende des Moduls vom ProjektStart entfernt ist</param>
    ''' <remarks></remarks>
    Public Sub moduleCopyTo(ByRef project As clsProjekt, ByVal parentID As String, ByVal moduleName As String, _
                                ByVal modulStartOffset As Integer, ByVal endOffset As Integer, ByVal dontStretch As Boolean)

        Dim moduleDauerInDays As Integer
        Dim phaseStartOffset As Integer = 0
        Dim correctFactor As Double
        Dim newphase As clsPhase
        Dim headPhase As clsPhase
        Dim elemID As String
        Dim parentPhase As clsPhase


        moduleDauerInDays = endOffset - modulStartOffset + 1
        correctFactor = moduleDauerInDays / Me.dauerInDays

        If correctFactor > 1.0 And dontStretch Then
            correctFactor = 1.0
        End If

        ' jetzt muss evtl eine Phase angelegt werden mit Namen moduleName, die dann die Sub-Phasen aufnimmt 
        ' das muss auf alle Fälle gemacht werden 
        ' danach ist auf alle Fälle in parentID die ID der Phase, die als Vater Phase dient
        If Not IsNothing(moduleName) Then

            If moduleName.Length > 0 Then
                headPhase = New clsPhase(parent:=project)
                elemID = project.hierarchy.findUniqueElemKey(moduleName, False)
                headPhase.nameID = elemID

                ' Änderung tk 17.11.15: die Phase 0 Ressourcen und Kosten übernehmen ..
                AllPhases.Item(0).korrCopyTo(headPhase, correctFactor, elemID)

                headPhase.changeStartandDauer(modulStartOffset, CLng(Me.dauerInDays * correctFactor))
                parentPhase = project.getPhaseByID(parentID)
                ' jetzt werden die Earliest und latest Spielräume für die Headphase, dann für die einzelnen Module eingetragen 

                headPhase.earliestStart = parentPhase.startOffsetinDays - headPhase.startOffsetinDays
                If headPhase.earliestStart > 0 Then
                    headPhase.earliestStart = 0
                End If

                headPhase.latestStart = parentPhase.startOffsetinDays + parentPhase.dauerInDays _
                                - (headPhase.startOffsetinDays + headPhase.dauerInDays)



                project.AddPhase(headPhase, origName:=moduleName, _
                       parentID:=parentID)

                parentID = elemID
            End If
        End If

        ' jetzt werden alle Phasen des Moduls übernommen
        Dim parentNameIDs(40) As String ' 41 Hierarchie-Stufen sollten genug sein 
        Dim currentLevel As Integer = 0
        Dim cphase As clsPhase
        Dim phaseID As String
        Dim tmpParentID As String
        ' die erste Phase ist ja die gesamte Modul-Länge, das ist ggf bereits im ersten Schritt erledigt worden   

        parentNameIDs(0) = parentID
        For p As Integer = 1 To Me.CountPhases - 1
            cphase = AllPhases.Item(p)
            currentLevel = Me.hierarchy.getIndentLevel(cphase.nameID)
            phaseID = project.hierarchy.findUniqueElemKey(cphase.name, False)
            parentNameIDs(currentLevel) = phaseID
            newphase = New clsPhase(project)

            ' die Namenszuweisung muss über diesen optionalen Parameter erfolgen . damit die Meilensteine richtig zugeordnet werden 
            AllPhases.Item(p).korrCopyTo(newphase, correctFactor, phaseID)
            ' jetzt muss diese Phase entsprechend im Projekt positioniert werden 

            phaseStartOffset = modulStartOffset + newphase.startOffsetinDays
            newphase.changeStartandDauer(phaseStartOffset, newphase.dauerInDays)

            If currentLevel - 1 < 0 Then
                tmpParentID = parentID
            Else
                tmpParentID = parentNameIDs(currentLevel - 1)
            End If

            parentPhase = project.getPhaseByID(tmpParentID)
            ' jetzt werden die Earliest und latest Spielräume für die Headphase, dann für die einzelnen Module eingetragen 

            newphase.earliestStart = parentPhase.startOffsetinDays - newphase.startOffsetinDays
            If newphase.earliestStart > 0 Then
                newphase.earliestStart = 0
            End If

            Try
                newphase.latestStart = parentPhase.startOffsetinDays + parentPhase.dauerInDays _
                            - (newphase.startOffsetinDays + newphase.dauerInDays)
            Catch ex As Exception
                Dim a = 2
            End Try




            project.AddPhase(phase:=newphase, origName:="", parentID:=tmpParentID)
        Next p

    End Sub

    ''' <summary>
    ''' gibt true zurück, wenn in der Vorlage irgendeiner der Meilensteine, entweder über BreadCrumb oder nur als Name angegeben, vorhanden ist
    ''' </summary>
    ''' <param name="msCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property containsAnyMilestonesOfCollection(ByVal msCollection As Collection) As Boolean
        Get
            Dim ix As Integer = 1
            Dim fullName As String
            Dim tmpResult As Boolean = False
            Dim containsMS As Boolean = False
            Dim tmpMilestone As clsMeilenstein

            If msCollection.Count = 0 Then
                tmpResult = True
            Else
                While ix <= msCollection.Count And Not containsMS

                    fullName = CStr(msCollection.Item(ix))
                    Dim curMsName As String = ""
                    Dim breadcrumb As String = ""
                    Dim pvName As String = ""
                    Dim type As Integer = -1

                    ' hier wird der Eintrag in filterMilestone aufgesplittet in curMsName und breadcrumb) 
                    Call splitHryFullnameTo2(fullName, curMsName, breadcrumb, type, pvName)

                    If type = -1 Or _
                        (type = PTProjektType.vorlage And pvName = Me.VorlagenName) Then

                        Dim milestoneIndices(,) As Integer = Me.hierarchy.getMilestoneIndices(curMsName, breadcrumb)
                        ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                        For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                            tmpMilestone = Me.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))
                            If IsNothing(tmpMilestone) Then

                            Else
                                containsMS = True
                                Exit For
                            End If

                        Next

                    End If

                    ix = ix + 1

                End While
                tmpResult = containsMS
            End If

            containsAnyMilestonesOfCollection = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn in der Vorlage irgendeiner der Meilensteine, entweder über BreadCrumb oder nur als Name angegeben, vorhanden ist
    ''' </summary>
    ''' <param name="phCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property containsAnyPhasesOfCollection(ByVal phCollection As Collection) As Boolean
        Get
            Dim ix As Integer = 1
            Dim fullName As String
            Dim tmpResult As Boolean = False
            Dim containsPH As Boolean = False
            Dim tmpPhase As clsPhase

            If phCollection.Count = 0 Then
                tmpResult = True
            Else
                While ix <= phCollection.Count And Not containsPH

                    fullName = CStr(phCollection.Item(ix))
                    Dim curPhName As String = ""
                    Dim breadcrumb As String = ""
                    Dim pvName As String = ""
                    Dim type As Integer = -1

                    ' hier wird der Eintrag in filterMilestone aufgesplittet in curMsName und breadcrumb) 
                    Call splitHryFullnameTo2(fullName, curPhName, breadcrumb, type, pvName)

                    If type = -1 Or _
                        (type = PTProjektType.vorlage And pvName = Me.VorlagenName) Then

                        Dim phaseIndices() As Integer = Me.hierarchy.getPhaseIndices(curPhName, breadcrumb)
                        ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                        For mx As Integer = 0 To CInt(phaseIndices.Length) - 1

                            tmpPhase = Me.getPhase(phaseIndices(mx))
                            If IsNothing(tmpPhase) Then

                            Else
                                containsPH = True
                                Exit For
                            End If

                        Next

                    End If

                    ix = ix + 1

                End While
                tmpResult = containsPH
            End If

            containsAnyPhasesOfCollection = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' kopiert die Hierarchie des aktuellen Me Projektes 
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <remarks></remarks>
    Friend Sub copyHryTo(ByRef newproject As clsProjekt)
        Dim ix As Integer
        Dim curNode As clsHierarchyNode
        Dim copiedNode As clsHierarchyNode
        Dim key As String
        Dim childKey As String

        newproject.hierarchy = New clsHierarchy

        For ix = 1 To Me.hierarchy.count
            curNode = Me.hierarchy.nodeItem(ix)
            key = Me.hierarchy.getIDAtIndex(ix)
            copiedNode = New clsHierarchyNode
            With copiedNode
                .elemName = curNode.elemName
                .indexOfElem = curNode.indexOfElem
                '.origName = curNode.origName
                .parentNodeKey = curNode.parentNodeKey
                ' jetzt die Kinder kopieren 
                For cx As Integer = 1 To curNode.childCount
                    childKey = curNode.getChild(cx)
                    .addChild(childKey)
                Next
            End With

            newproject.hierarchy.copyNode(copiedNode, key)

        Next

    End Sub


    Public ReadOnly Property Liste() As List(Of clsPhase)

        Get
            Liste = AllPhases
        End Get

    End Property

    Public Overridable ReadOnly Property dauerInDays As Integer

        Get
            Dim i As Integer
            Dim max As Double = 0

            ' Bestimmung der Dauer 

            For i = 1 To Me.CountPhases

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays Then
                        max = .startOffsetinDays + .dauerInDays
                    End If

                End With

            Next i


            dauerInDays = CInt(max)
            _Dauer = getColumnOfDate(StartofCalendar.AddDays(max - 1))

        End Get
    End Property


    Public ReadOnly Property anzahlRasterElemente() As Integer


        Get

            Dim tmpValue As Integer = 0

            If Me.CountPhases > 0 Then
                With Me.getPhase(1)
                    tmpValue = .relEnde - .relStart + 1
                End With
            End If

            anzahlRasterElemente = tmpValue


        End Get

    End Property

    Public Property UID() As Long

        Get
            UID = uuid
        End Get

        Set(value As Long)
            uuid = value
        End Set

    End Property

    Public ReadOnly Property CountPhases() As Integer

        Get
            CountPhases = AllPhases.Count
        End Get

    End Property

    ''' <summary>
    ''' gibt die Phase mit Index zurück, wenn Index kleiner bzw. gleich 1 oder größer Anzahl Phasen, 
    ''' dann Nothing 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhase(ByVal index As Integer) As clsPhase

        Get
            If index < 1 Or index > AllPhases.Count Then
                getPhase = Nothing
            Else
                getPhase = AllPhases.Item(index - 1)
            End If
        End Get
    End Property

    Public ReadOnly Property getPhaseCount(ByVal name As String, Optional ByVal breadcrumb As String = "") As Integer
        Get

            Dim phaseIndices() As Integer = Me.hierarchy.getPhaseIndices(name, breadcrumb)
            If phaseIndices.Length = 1 And phaseIndices(0) = 0 Then
                getPhaseCount = 0
            Else
                getPhaseCount = phaseIndices.Length
            End If

        End Get
    End Property

    

    ''' <summary>
    ''' gibt zurück, ob die Parent-Phase mit ID=parentID identisch zur Phase mit Name elemName, startdate, endDate ist) 
    ''' </summary>
    ''' <param name="elemName">Name der Phase</param>
    ''' <param name="startDate">Start-Datum der Phase</param>
    ''' <param name="endDate">Ende-Datum der Phase</param>
    ''' <param name="tolerance" >tolerance = 1.0: bedeutet, daß 100% üÜbereinstimmung vorliegen muss</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isCloneToParent(ByVal elemName As String, ByVal parentID As String, ByVal startDate As Date, ByVal endDate As Date, _
                                         Optional ByVal tolerance As Double = 1.0) As Boolean

        Dim parentPhase As clsPhase = Me.getPhaseByID(parentID)
        Dim istIdentisch As Boolean = False

        If Not IsNothing(parentPhase) Then

            Dim parentStartDate As Date = parentPhase.getStartDate
            Dim parentEndDate As Date = parentPhase.getEndDate

            Dim ueberdeckung As Double = calcPhaseUeberdeckung(parentStartDate, parentEndDate, _
                                                                startDate, endDate)

            If elemNameOfElemID(parentID) = elemName And ueberdeckung >= tolerance Then
                istIdentisch = True
            Else
                istIdentisch = False
            End If


        Else
            istIdentisch = False
        End If

        isCloneToParent = istIdentisch

    End Function

    ''' <summary>
    ''' gibt die ID des Sibling Elements zurück, das von Namen, Start , Ausdehnung innerhalb der Toleranz identisch ist 
    ''' leerer String, wenn das Element nicht existiert
    ''' </summary>
    ''' <param name="parentID">ID in der Hierachy vom Parent-Knoten</param>
    ''' <param name="elemName">Name der Phase</param>
    ''' <param name="startDate">Start-Datum der Phase</param>
    ''' <param name="endDate">Ende-Datum der Phase</param>
    ''' <param name="tolerance">Toleranz, innerhalb der die Überdeckung als identisch gilt; ohne Angabe wird nur 100% Überdeckung als identisch angesehen</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDuplicatePhaseSiblingID(ByVal elemName As String, ByVal parentID As String, ByVal startDate As Date, ByVal endDate As Date, _
                                             Optional ByVal tolerance As Double = 1.0) As String

        Dim parentNode As clsHierarchyNode = Me.hierarchy.nodeItem(parentID)
        Dim siblingID As String
        Dim identicalID As String = ""

        Dim anzahlKinder As Integer = parentNode.childCount
        Dim istIdentisch As Boolean = False
        Dim currentChildNr As Integer = 1

        Dim siblingPhase As clsPhase = Me.getPhaseByID(parentID)

        Do While Not istIdentisch And currentChildNr <= anzahlKinder

            siblingID = parentNode.getChild(currentChildNr)

            If Not elemIDIstMeilenstein(siblingID) Then

                siblingPhase = Me.getPhaseByID(siblingID)

                If Not IsNothing(siblingPhase) Then

                    Dim siblingStartDate As Date = siblingPhase.getStartDate
                    Dim siblingEndDate As Date = siblingPhase.getEndDate

                    Dim ueberdeckung As Double = calcPhaseUeberdeckung(siblingStartDate, siblingEndDate, _
                                                                        startDate, endDate)

                    If siblingPhase.name = elemName And ueberdeckung >= tolerance Then
                        istIdentisch = True
                        identicalID = siblingID
                    Else
                        istIdentisch = False
                    End If


                Else
                    istIdentisch = False
                End If

            End If

            currentChildNr = currentChildNr + 1

        Loop


        getDuplicatePhaseSiblingID = identicalID

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="parentID"></param>
    ''' <param name="msName"></param>
    ''' <param name="msDate"></param>
    ''' <param name="toleranceInDays"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getDuplicateMsSiblingID(ByVal msName As String, ByVal parentID As String, ByVal msDate As Date, _
                                                  Optional ByVal toleranceInDays As Integer = 0) As String

        Dim parentNode As clsHierarchyNode = Me.hierarchy.nodeItem(parentID)
        Dim siblingID As String
        Dim identicalID As String = ""

        Dim anzahlKinder As Integer = parentNode.childCount
        Dim istIdentisch As Boolean = False
        Dim currentChildNr As Integer = 1

        Dim siblingMilestone As clsMeilenstein

        Do While Not istIdentisch And currentChildNr <= anzahlKinder

            siblingID = parentNode.getChild(currentChildNr)

            If elemIDIstMeilenstein(siblingID) Then

                siblingMilestone = Me.getMilestoneByID(siblingID)

                If Not IsNothing(siblingMilestone) Then

                    Dim siblingDate As Date = siblingMilestone.getDate

                    Dim diffInDays = DateDiff(DateInterval.Day, siblingDate, msDate)
                    If diffInDays < 0 Then
                        diffInDays = diffInDays * -1
                    End If

                    If diffInDays <= toleranceInDays And siblingMilestone.name = msName Then
                        istIdentisch = True
                        identicalID = siblingMilestone.nameID
                    Else
                        istIdentisch = False
                    End If

                Else
                    istIdentisch = False
                End If

            End If

            currentChildNr = currentChildNr + 1

        Loop


        getDuplicateMsSiblingID = identicalID


    End Function

    ''' <summary>
    ''' gibt true zurück, wenn es sich um eine Phase der Gliederungsebene 1 handelt, also Kind-Phase der rootphase ist
    ''' gibt false sonst zurück
    ''' wenn BHTC Schema = true, dann muss es ein Kind der ersten oder zweiten Hierarchie Ebene handeln   
    ''' </summary>
    ''' <param name="elemName">Name der Phase</param>
    ''' <param name="isBHTCSchema"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isSwimlaneOrSegment(ByVal elemName As String, Optional ByVal isBHTCSchema As Boolean = False) As Boolean
        Get
            Dim tmpResult As Boolean = False
            Dim ChildCollection As Collection = Me.hierarchy.getChildIDsOf(rootPhaseName, False)
            Dim itemNameID As String
            Dim childNameID As String
            Dim fullBC As String
            Dim fullChildBC As String

            ' wenn die RootPhase Meilensteine enthält, dann ist sie eine Swimlane bzw ein Segment 
            If elemName = "." And Me.hierarchy.getChildIDsOf(rootPhaseName, True).Count > 0 Then
                tmpResult = True
            Else
                elemName = elemName & "#"

                ' muss auch true zurück geben, wenn es sich um die rootPhase handelt und Meilensteine drin vorkommen 


                If isBHTCSchema Then
                    ' noch nicht implementiert 
                    Dim found As Boolean = False
                    Dim ix As Integer = 1

                    Do While ix <= ChildCollection.Count And Not found
                        itemNameID = CStr(ChildCollection.Item(ix))
                        fullBC = Me.getBcElemName(itemNameID)

                        If fullBC.EndsWith(elemName) Then
                            tmpResult = True
                            found = True
                        Else
                            If Not elemIDIstMeilenstein(itemNameID) Then
                                Dim childChildCollection As Collection = Me.hierarchy.getChildIDsOf(itemNameID, False)
                                ' Schleife über das KindesKin
                                For Each childNameID In childChildCollection
                                    fullChildBC = Me.getBcElemName(childNameID)
                                    If fullChildBC.EndsWith(elemName) Then
                                        tmpResult = True
                                        found = True
                                        Exit For
                                    End If
                                Next

                            End If

                        End If
                        ix = ix + 1
                    Loop


                Else

                    For Each itemNameID In ChildCollection
                        fullBC = Me.getBcElemName(itemNameID)
                        If fullBC.EndsWith(elemName) Then
                            tmpResult = True
                            Exit For
                        End If
                    Next

                End If
            End If

            isSwimlaneOrSegment = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl an Meilenstein bzw. Phasen Kategorien zurück 
    ''' </summary>
    ''' <param name="lookingforMS"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCategoryCount(ByVal lookingforMS As Boolean) As Integer
        Get
            Dim anzCategories As Integer = 0
            Dim nameCollection As New Collection
            Dim idCollection As Collection = Me.getAllElemIDs(lookingforMS)

            For Each elemID As String In idCollection
                If lookingforMS Then
                    Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(elemID)
                    If Not IsNothing(cMilestone) Then
                        Dim catName As String = cMilestone.appearance
                        If Not nameCollection.Contains(catName) Then
                            nameCollection.Add(Item:=catName, Key:=catName)
                        End If
                    End If
                Else
                    Dim cPhase As clsPhase = Me.getPhaseByID(elemID)
                    If Not IsNothing(cPhase) Then
                        Dim catName As String = cPhase.appearance
                        If Not nameCollection.Contains(catName) Then
                            nameCollection.Add(Item:=catName, Key:=catName)
                        End If
                    End If
                End If
                
            Next

            getCategoryCount = nameCollection.Count

        End Get
    End Property

    ''' <summary>
    ''' gibt die Namen der vorkommenden Meilenstein- bzw. Phasen Kategorien zurück 
    ''' </summary>
    ''' <param name="lookingforMS"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCategoryNames(ByVal lookingforMS As Boolean) As Collection
        Get
            Dim anzCategories As Integer = 0
            Dim nameCollection As New Collection
            Dim idCollection As Collection = Me.getAllElemIDs(lookingforMS)
            Dim tmpSortList As New SortedList(Of Integer, String)

            For Each elemID As String In idCollection
                If lookingforMS Then
                    Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(elemID)
                    If Not IsNothing(cMilestone) Then
                        Dim catName As String = cMilestone.appearance
                        Dim sortkey As Integer = appearanceDefinitions.IndexOfKey(catName)
                        If Not tmpSortList.ContainsKey(sortkey) Then
                            tmpSortList.Add(key:=sortkey, value:=catName)
                        End If
                    End If
                Else
                    Dim cPhase As clsPhase = Me.getPhaseByID(elemID)
                    If Not IsNothing(cPhase) Then
                        Dim catName As String = cPhase.appearance
                        Dim sortkey As Integer = appearanceDefinitions.IndexOfKey(catName)
                        If Not tmpSortList.ContainsKey(sortkey) Then
                            tmpSortList.Add(key:=sortkey, value:=catName)
                        End If
                    End If
                End If

            Next

            ' jetzt umkopieren 

            For i As Integer = 0 To tmpSortList.Count - 1
                Dim catName As String = tmpSortList.ElementAt(i).Value
                If Not nameCollection.Contains(catName) Then
                    nameCollection.Add(catName, catName)
                End If
            Next

            getCategoryNames = nameCollection
        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl der Swimlanes zurück, die für das Projekt bei der gegebenen Menge von Phasen und Meilensteinen gezeichnet werden müssen; 
    ''' dabei wird unterschieden, ob es sich um das BHTC Schema handelt oder um eine freie Swimlane Definition handelt  
    ''' </summary>
    ''' <param name="considerAll">sollen alle Swimlanes betrachtet werden</param>
    ''' <param name="breadCrumbArray">enthält eine Liste der vollständigen Breadcrumbs aller gewählten Elemente </param>
    ''' <param name="isBhtcSchema">gibt an, ob die Swimlanes dadurch bestimmt sind, dass sie auf der 2. Ebene sind oder ob alles als Swimlane behandelt wird, was erst bei der 2.Stufe (wie bei BHTC) 
    ''' in einer ersten Implementierung wird Level2 betrachtet</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSwimLanesCount(ByVal considerAll As Boolean, _
                                               ByVal breadCrumbArray() As String, _
                                               ByVal isBhtcSchema As Boolean) As Integer
        Get
            Dim anzSwimlanes As Integer = 0
            Dim parentNameID As String = rootPhaseName

            Dim sptr As Integer
            ' ein fullBreadCrumb enthält auch den ElemName am Schluss ...
            Dim fullSwlBreadCrumb As String

            If IsNothing(breadCrumbArray) Then
                sptr = -1
            Else
                sptr = breadCrumbArray.Length - 1
            End If

            If Not isBhtcSchema Then

                ' gibt es Meilensteine in der Rootphase? 
                Dim anzMilestonesInRootPhase As Integer = Me.hierarchy.getChildIDsOf(rootPhaseName, True).Count
                If anzMilestonesInRootPhase > 0 Then
                    fullSwlBreadCrumb = Me.getBcElemName(rootPhaseName)
                    If considerAll Then
                        anzSwimlanes = anzSwimlanes + 1
                    Else
                        fullSwlBreadCrumb = Me.getBcElemName(rootPhaseName)
                        ' ist eines der Elemente in der aktuellen Swimlane enthalten ? 
                        Dim found As Boolean = False
                        Dim index = 0
                        Do While Not found And index <= sptr
                            If breadCrumbArray(index).StartsWith(fullSwlBreadCrumb) Then
                                found = True
                            Else
                                index = index + 1
                            End If
                        Loop
                        If found Then
                            anzSwimlanes = anzSwimlanes + 1
                        End If
                    End If
                End If

                ' jetzt kommen die Phasen Kinder der RootPhase  
                Dim ChildCollection As Collection = Me.hierarchy.getChildIDsOf(rootPhaseName, False)

                For Each childObj As Object In ChildCollection
                    Dim swimlaneID As String = CStr(childObj)
                    fullSwlBreadCrumb = Me.getBcElemName(swimlaneID)

                    If considerAll Then
                        anzSwimlanes = anzSwimlanes + 1
                    Else
                        ' ist eines der Elemente in der aktuellen Swimlane enthalten ? 
                        Dim found As Boolean = False
                        Dim index = 0
                        Do While Not found And index <= sptr
                            If breadCrumbArray(index).StartsWith(fullSwlBreadCrumb) Then
                                found = True
                            Else
                                index = index + 1
                            End If
                        Loop
                        If found Then
                            anzSwimlanes = anzSwimlanes + 1
                        End If
                    End If
                    ' wenn jetzt die Auswahl = 0 ist, dann sollen alle betrachtet werden ...

                Next

            Else
                Dim ankerName As String = "BHTC milestones"
                Dim ankerPhase As clsPhase = Me.getPhase(ankerName)


                If Not IsNothing(ankerPhase) Then
                    parentNameID = Me.hierarchy.getParentIDOfID(ankerPhase.nameID)
                End If

                If parentNameID.Length > 1 Then

                    Dim segmentCollection As Collection = Me.hierarchy.getChildIDsOf(parentNameID, False)

                    For Each obj As Object In segmentCollection
                        Dim sgementNameID As String = CStr(obj)
                        Dim ChildCollection As Collection = Me.hierarchy.getChildIDsOf(sgementNameID, False)

                        For Each childObj As Object In ChildCollection
                            Dim swimlaneID As String = CStr(childObj)
                            fullSwlBreadCrumb = Me.getBcElemName(swimlaneID)

                            If considerAll Then
                                anzSwimlanes = anzSwimlanes + 1
                            Else
                                ' ist eines der Elemente in der aktuellen Swimlane enthalten ? 
                                Dim found As Boolean = False
                                Dim index = 0
                                Do While Not found And index <= sptr
                                    If breadCrumbArray(index).StartsWith(fullSwlBreadCrumb) Then
                                        found = True
                                    Else
                                        index = index + 1
                                    End If
                                Loop
                                If found Then
                                    anzSwimlanes = anzSwimlanes + 1
                                End If
                            End If
                            ' wenn jetzt die Auswahl = 0 ist, dann sollen alle betrachtet werden ...

                        Next

                    Next

                End If



            End If

            getSwimLanesCount = anzSwimlanes

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl der Segmente, das heisst die Anzahl Phasen auf Hierarchie-Stufe 1 zurück 
    ''' </summary>
    ''' <param name="considerAll">sollen alle Elemente betrachtet werden </param>
    ''' <param name="breadCrumbArray">es sollen nur die Elemente betrachtet werden, die </param>
    ''' <param name="isBhtcSchema"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSegmentsCount(ByVal considerAll As Boolean, _
                                                   ByVal breadCrumbArray() As String, _
                                                   ByVal isBhtcSchema As Boolean) As Integer
        Get

            Dim ankerName As String = "BHTC milestones"
            Dim anzSegments = 0
            Dim sptr As Integer

            If Not IsNothing(breadCrumbArray) Then
                sptr = breadCrumbArray.Length - 1
            Else
                sptr = -1
            End If

            Dim fullSwlBreadCrumb As String


            If Not isBhtcSchema Then
                anzSegments = 0
            Else
                Dim ankerPhase As clsPhase = Me.getPhase(ankerName)
                Dim parentNameID As String = ""

                If Not IsNothing(ankerPhase) Then
                    parentNameID = Me.hierarchy.getParentIDOfID(ankerPhase.nameID)
                End If

                If parentNameID.Length > 1 Then

                    Dim segmentCollection As Collection = Me.hierarchy.getChildIDsOf(rootPhaseName, False)

                    For Each obj As Object In segmentCollection
                        Dim segmentNameID As String = CStr(obj)
                        fullSwlBreadCrumb = Me.getBcElemName(segmentNameID)

                        If considerAll Then
                            anzSegments = anzSegments + 1
                        Else
                            ' ist eines der Elemente im aktuellen Segment enthalten ? 
                            Dim found As Boolean = False
                            Dim index = 0
                            Do While Not found And index <= sptr
                                If breadCrumbArray(index).StartsWith(fullSwlBreadCrumb) Then
                                    found = True
                                Else
                                    index = index + 1
                                End If
                            Loop
                            If found Then
                                anzSegments = anzSegments + 1
                            End If
                        End If

                    Next

                End If

            End If

            getSegmentsCount = anzSegments

        End Get
    End Property


    ''' <summary>
    ''' gibt die Swimlane mit der Reihenfolge-Nr "index" zurück; index läuft von 1..Anzahl 
    ''' Nothing, wenn es die entsprechende Swimlane gar nicht gibt  
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="considerAll">sollen alle betrachtet werden</param>
    ''' <param name="breadCrumbArray">enthält die Breadcrumbs der zu betrachtenden Swimlanes</param>
    ''' <param name="isBhtcSchema"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSwimlane(ByVal index As Integer, ByVal considerAll As Boolean, _
                                               ByVal breadCrumbArray() As String, _
                                               ByVal isBhtcSchema As Boolean) As clsPhase
        Get

            Dim tmpPhase As clsPhase = Nothing
            Dim ankerName As String = "BHTC milestones"
            Dim parentNameID As String = rootPhaseName

            Dim anzSwimlanes As Integer = 0
            Dim sptr As Integer = -1

            If Not IsNothing(breadCrumbArray) Then
                sptr = breadCrumbArray.Length - 1
            End If



            ' ein fullBreadCrumb enthält auch den ElemName am Schluss ...
            Dim fullSwlBreadCrumb As String

            If index > 0 Then

                If Not isBhtcSchema Then

                    ' gibt es Meilensteine in der Rootphase? 
                    Dim anzMilestonesInRootPhase As Integer = Me.hierarchy.getChildIDsOf(rootPhaseName, True).Count
                    If anzMilestonesInRootPhase > 0 Then
                        fullSwlBreadCrumb = Me.getBcElemName(rootPhaseName)
                        If considerAll Then
                            anzSwimlanes = anzSwimlanes + 1
                            If index = anzSwimlanes Then
                                ' das ist jetzt die Phase 
                                tmpPhase = Me.getPhaseByID(rootPhaseName)
                            End If
                        Else
                            fullSwlBreadCrumb = Me.getBcElemName(rootPhaseName)
                            ' ist eines der Elemente in der aktuellen Swimlane enthalten ? 
                            Dim found As Boolean = False
                            Do While Not found And index <= sptr
                                If breadCrumbArray(index).StartsWith(fullSwlBreadCrumb) Then
                                    found = True
                                Else
                                    index = index + 1
                                End If
                            Loop
                            If found Then
                                anzSwimlanes = anzSwimlanes + 1
                            End If
                        End If
                    End If

                    ' das jetzt nur machen, wenn tmpPhase noch immer Nothing ist ... 

                    If IsNothing(tmpPhase) Then

                        Dim ChildCollection As Collection = Me.hierarchy.getChildIDsOf(rootPhaseName, False)

                        For Each childObj As Object In ChildCollection
                            Dim swimlaneID As String = CStr(childObj)
                            fullSwlBreadCrumb = Me.getBcElemName(swimlaneID)

                            If considerAll Then
                                anzSwimlanes = anzSwimlanes + 1
                                If index = anzSwimlanes Then
                                    ' das ist jetzt die Phase 
                                    tmpPhase = Me.getPhaseByID(swimlaneID)
                                    Exit For
                                End If
                            Else
                                ' ist eines der Elemente in der Swimlane enthalten ? 
                                Dim found As Boolean = False
                                Dim ix = 0
                                Do While Not found And ix <= sptr
                                    If breadCrumbArray(ix).StartsWith(fullSwlBreadCrumb) Then
                                        found = True
                                    Else
                                        ix = ix + 1
                                    End If
                                Loop
                                If found Then
                                    anzSwimlanes = anzSwimlanes + 1
                                    If index = anzSwimlanes Then
                                        ' das ist jetzt die Phase 
                                        tmpPhase = Me.getPhaseByID(swimlaneID)
                                        Exit For
                                    End If
                                End If
                            End If
                            ' wenn jetzt die Auswahl = 0 ist, dann sollen alle betrachtet werden ...

                        Next

                    End If
                    

                Else
                    Dim ankerPhase As clsPhase = Me.getPhase(ankerName)

                    If Not IsNothing(ankerPhase) Then
                        parentNameID = Me.hierarchy.getParentIDOfID(ankerPhase.nameID)
                    End If

                    If parentNameID.Length > 1 Then

                        Dim segmentCollection As Collection = Me.hierarchy.getChildIDsOf(parentNameID, False)

                        For Each obj As Object In segmentCollection
                            Dim segmentNameID As String = CStr(obj)
                            Dim ChildCollection As Collection = Me.hierarchy.getChildIDsOf(segmentNameID, False)

                            For Each childObj As Object In ChildCollection
                                Dim swimlaneID As String = CStr(childObj)
                                fullSwlBreadCrumb = Me.getBcElemName(swimlaneID)

                                If considerAll Then
                                    anzSwimlanes = anzSwimlanes + 1
                                    If index = anzSwimlanes Then
                                        ' das ist jetzt die Phase 
                                        tmpPhase = Me.getPhaseByID(swimlaneID)
                                        Exit For
                                    End If
                                Else
                                    ' ist eines der Elemente in der Swimlane enthalten ? 
                                    Dim found As Boolean = False
                                    Dim ix = 0
                                    Do While Not found And ix <= sptr
                                        If breadCrumbArray(ix).StartsWith(fullSwlBreadCrumb) Then
                                            found = True
                                        Else
                                            ix = ix + 1
                                        End If
                                    Loop
                                    If found Then
                                        anzSwimlanes = anzSwimlanes + 1
                                        If index = anzSwimlanes Then
                                            ' das ist jetzt die Phase 
                                            tmpPhase = Me.getPhaseByID(swimlaneID)
                                            Exit For
                                        End If
                                    End If
                                End If
                                ' wenn jetzt die Auswahl = 0 ist, dann sollen alle betrachtet werden ...

                            Next

                            If index = anzSwimlanes Then
                                Exit For
                            End If

                        Next

                    End If


                End If

            End If

            getSwimlane = tmpPhase

        End Get
    End Property

    ''' <summary>
    ''' gibt für eine übergebene ID den vollen BreadCrum+ElemName zurück
    ''' den Abschluss bildet stets ein #, damit über ..startwith festgestellt werden kann, ob ein Element Kind eines anderen ist 
    ''' ?, wenn PhaseID nicht existiert 
    ''' </summary>
    ''' <param name="nameID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBcElemName(ByVal nameID As String) As String
        Get

            Dim elemName As String = ""
            Dim fullName As String = "?"

            If istElemID(nameID) Then
                elemName = elemNameOfElemID(nameID)
                fullName = calcHryFullname(elemName, _
                                              Me.hierarchy.getBreadCrumb(nameID)) & "#"
            End If

            getBcElemName = fullName

        End Get
    End Property


    ''' <summary>
    ''' liefert einen Array zurück, der die Breadcrumbs inkl. der Namen der in selectedPhaseIDs und selectedMilestoneIDs übergebenen 
    ''' Elemente enthält 
    ''' </summary>
    ''' <param name="selectedPhaseIDs">Liste der PhaseIDs</param>
    ''' <param name="selectedMilestoneIDs">Liste der MilestoneIDs</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBreadCrumbArray(ByVal selectedPhaseIDs As Collection, ByVal selectedMilestoneIDs As Collection) As String()
        Get
            ' im Array suchfeld sind die 
            Dim suchDimension As Integer = selectedPhaseIDs.Count + selectedMilestoneIDs.Count - 1


            If suchDimension >= 0 Then
                Dim suchFeld(suchDimension) As String
                Dim elemID As String = ""
                Dim sptr As Integer = 0

                ' suchfeld aufbauen 
                For i As Integer = 1 To selectedPhaseIDs.Count
                    elemID = CStr(selectedPhaseIDs.Item(i))
                    suchFeld(sptr) = Me.getBcElemName(elemID)
                    sptr = sptr + 1
                Next

                For i As Integer = 1 To selectedMilestoneIDs.Count
                    elemID = CStr(selectedMilestoneIDs.Item(i))
                    suchFeld(sptr) = Me.getBcElemName(elemID)
                    sptr = sptr + 1
                Next

                sptr = sptr - 1
                getBreadCrumbArray = suchFeld

            Else
                getBreadCrumbArray = Nothing
            End If


        End Get
    End Property


    ''' <summary>
    ''' gibt alle Phasen bzw. Milestone ElemIDs in einer Collection zurück 
    ''' die Milestones gehen alle mit 1§ los, die Phasen alle mit 0§; 
    ''' deshalb markiert das erste Element mit "1§" das Ende der Phasen bzw. 
    ''' den Start der Meilensteine
    ''' </summary>
    ''' <param name="lookingForMS"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllElemIDs(ByVal lookingForMS As Boolean) As Collection
        Get
            Dim iDCollection As New Collection
            Dim tmpSortList As New SortedList(Of DateTime, String)
            Dim sortDate As DateTime
            Dim firstIX As Integer, lastIX As Integer
            Dim elemID As String

            If lookingForMS Then
                lastIX = Me.hierarchy.count
                firstIX = Me.hierarchy.getIndexOf1stMilestone
                If firstIX < 0 Then
                    ' es gibt keine Meilensteine 
                Else

                    For mx = firstIX To lastIX
                        elemID = Me.hierarchy.getIDAtIndex(mx)

                        sortDate = Me.getMilestoneByID(elemID).getDate
                        If Not tmpSortList.ContainsValue(elemID) Then

                            Do While tmpSortList.ContainsKey(sortDate)
                                sortDate = sortDate.AddMilliseconds(1)
                            Loop

                            tmpSortList.Add(sortDate, elemID)

                        End If

                    Next

                End If
            Else
                ' Phasen holen
                firstIX = 1
                lastIX = Me.hierarchy.getIndexOf1stMilestone - 1

                If lastIX < 0 Then
                    ' es gibt keine Meilensteine, sondern nur Phasen 
                    lastIX = Me.hierarchy.count
                End If

                For mx = firstIX To lastIX
                    elemID = Me.hierarchy.getIDAtIndex(mx)

                    sortDate = Me.getPhaseByID(elemID).getStartDate

                    If Not tmpSortList.ContainsValue(elemID) Then

                        Do While tmpSortList.ContainsKey(sortDate)
                            sortDate = sortDate.AddMilliseconds(1)
                        Loop


                        tmpSortList.Add(sortDate, elemID)


                    End If

                Next

            End If

            ' jetzt muss umkopiert werden 
            For Each kvp As KeyValuePair(Of DateTime, String) In tmpSortList
                iDCollection.Add(kvp.Value, kvp.Value)
            Next

            getAllElemIDs = iDCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt die Phase zurück, die die folgenden Eigenschaften erfüllt
    ''' hat name als elemName Bestandteil
    ''' hat den optional angegebenen Breadcrumb, wenn der nicht angegeben wird oder "" ist, dann ist es egal, unter welcher Hierarchie Stufe die Phase liegen soll 
    ''' der breadcrum kann die gesamte Hierarchie umfassen oder auch nur die erste Parent-Stufe; Parent-Stufen werden per # voneinander getrennt
    ''' hat die optional angegebene lfdNr, ist also das lfdNr-vielte Vorkommen von name / breadcrumb 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <param name="lfdNr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhase(ByVal name As String, Optional ByVal breadcrumb As String = "", Optional ByVal lfdNr As Integer = 1) As clsPhase

        Get

            Dim found As Boolean = False


            Dim phaseIndices() As Integer
            phaseIndices = Me.hierarchy.getPhaseIndices(name, breadcrumb)

            If lfdNr > phaseIndices.Length Or lfdNr < 1 Then
                getPhase = Nothing
            Else
                If phaseIndices(lfdNr - 1) > 0 And phaseIndices(lfdNr - 1) <= AllPhases.Count Then
                    ' wenn phaseIndices(x) = 0 dann gibt es die Phase nicht ..
                    getPhase = AllPhases.Item(phaseIndices(lfdNr - 1) - 1)
                Else
                    getPhase = Nothing
                End If

            End If



            ' alter Code
            'found = False
            'i = 1
            'While i <= AllPhases.Count And Not found
            '    If name = AllPhases.Item(i - 1).name Then
            '        found = True
            '        index = i
            '    Else
            '        i = i + 1
            '    End If

            'End While

            'If found Then
            '    getPhase = AllPhases.Item(index - 1)
            'Else
            '    getPhase = Nothing
            'End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die der ElemID entsprechende Phase zurück 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseByID(ByVal elemID As String) As clsPhase

        Get

            Dim phIndex As Integer = Me.hierarchy.getPMIndexOfID(elemID)
            If phIndex >= 0 Or phIndex < AllPhases.Count Then
                getPhaseByID = AllPhases.Item(phIndex - 1)
            Else
                getPhaseByID = Nothing
            End If


        End Get

    End Property

    '
    ' übergibt in getPhasenBedarf die Werte der Phase <phaseid>
    '
    Public Overridable ReadOnly Property getPhasenBedarf(phaseName As String) As Double()

        Get
            Dim phaseValues() As Double
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer
            Dim phase As clsPhase


            If _Dauer > 0 Then

                ReDim phaseValues(_Dauer - 1)

                anzPhasen = AllPhases.Count
                If anzPhasen > 0 Then

                    For p = 0 To anzPhasen - 1
                        phase = AllPhases.Item(p)

                        If phase.name = phaseName Then
                            With phase
                                For i = .relStart To .relEnde
                                    phaseValues(i - 1) = phaseValues(i - 1) + 1
                                Next
                            End With

                        End If

                    Next p ' Loop über alle Phasen
                Else
                    Throw New ArgumentException("Projekt hat keine Phasen")
                End If


                getPhasenBedarf = phaseValues

            Else
                Throw New ArgumentException("Projekt hat keine Dauer")
                getPhasenBedarf = phaseValues
            End If
        End Get

    End Property

    '
    ' übergibt in getRessourcenBedarf die Werte der Rolle <roleid>
    '
    Public ReadOnly Property getRessourcenBedarf(roleID As Object) As Double()

        Get
            Dim roleValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer
            Dim found As Boolean
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim lookforIndex As Boolean
            Dim phasenStart As Integer
            Dim tempArray As Double()


            If _Dauer > 0 Then

                lookforIndex = IsNumeric(roleID)

                ReDim roleValues(_Dauer - 1)

                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzRollen = .countRoles
                        phasenStart = .relStart - 1

                        ' Änderung: relende, relstart bezeichnet nicht mehr notwendigerweise die tatsächliche Länge des Arrays
                        ' es können Unschärfen auftreten 
                        'phasenEnde = .relEnde - 1


                        For r = 1 To anzRollen
                            role = .getRole(r)
                            found = False

                            With role
                                If lookforIndex Then
                                    If .RollenTyp = CInt(roleID) Then
                                        found = True
                                    End If
                                Else
                                    If .name = CStr(roleID) Then
                                        found = True
                                    End If
                                End If

                                Dim dimension As Integer
                                If found Then
                                    dimension = .getDimension
                                    ReDim tempArray(dimension)
                                    tempArray = .Xwerte
                                    For i = phasenStart To phasenStart + dimension
                                        roleValues(i) = roleValues(i) + tempArray(i - phasenStart)
                                    Next i
                                End If
                            End With ' role

                        Next r

                    End With ' phase


                Next p ' Loop über alle Phasen

                getRessourcenBedarf = roleValues

            Else
                ReDim roleValues(0)
                getRessourcenBedarf = roleValues
            End If
        End Get

    End Property

    Public ReadOnly Property getRessourcenBedarfNew(ByVal roleID As Object, Optional ByVal inclSubRoles As Boolean = False) As Double()

        Get
            Dim roleValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer
            'Dim found As Boolean
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim lookforIndex As Boolean
            Dim phasenStart As Integer
            Dim tempArray As Double()
            Dim roleUID As Integer
            Dim roleName As String = ""

            Dim roleIDs As New SortedList(Of Integer, Double)


            If _Dauer > 0 Then

                lookforIndex = IsNumeric(roleID)
                If IsNumeric(roleID) Then
                    roleUID = CInt(roleID)
                    roleName = RoleDefinitions.getRoledef(roleUID).name
                Else
                    If RoleDefinitions.containsName(CStr(roleID)) Then
                        roleUID = RoleDefinitions.getRoledef(CStr(roleID)).UID
                        roleName = CStr(roleID)
                    End If
                End If

                ' jetzt prüfen, ob es inkl aller SubRoles sein soll 
                If inclSubRoles Then
                    roleIDs = RoleDefinitions.getSubRoleIDsOf(roleName, type:=PTcbr.all)
                Else
                    roleIDs.Add(roleUID, 1.0)
                End If


                ReDim roleValues(_Dauer - 1)

                For Each srkvp As KeyValuePair(Of Integer, Double) In roleIDs
                    roleName = RoleDefinitions.getRoledef(srkvp.Key).name

                    Dim listOfPhases As Collection = Me.rcLists.getPhasesWithRole(roleName, False)
                    anzPhasen = listOfPhases.Count

                    For p = 1 To anzPhasen
                        phase = Me.getPhaseByID(CStr(listOfPhases.Item(p)))

                        With phase
                            ' Off1
                            anzRollen = .countRoles
                            phasenStart = .relStart - 1

                            ' Änderung: relende, relstart bezeichnet nicht mehr notwendigerweise die tatsächliche Länge des Arrays
                            ' es können Unschärfen auftreten 
                            'phasenEnde = .relEnde - 1


                            For r = 1 To anzRollen
                                role = .getRole(r)

                                With role

                                    If .RollenTyp = srkvp.Key Then
                                        Dim dimension As Integer

                                        dimension = .getDimension
                                        ReDim tempArray(dimension)
                                        tempArray = .Xwerte

                                        For i = phasenStart To phasenStart + dimension
                                            roleValues(i) = roleValues(i) + tempArray(i - phasenStart)
                                        Next i

                                    End If

                                End With ' role
                            Next r

                        End With ' phase


                    Next p ' Loop über alle Phasen

                Next ' Loop über for each srkvp

                getRessourcenBedarfNew = roleValues

            Else
                ReDim roleValues(0)
                getRessourcenBedarfNew = roleValues
            End If
        End Get

    End Property

    '
    ' übergibt in getRoleNames eine Collection von Rollen Definitionen, das sind alle Rollen, die in den Phasen vorkommen und einen Bedarf von größer Null haben
    '
    ''' <summary>
    ''' gibt die Liste aller im Projekt vergebenen Rollen aus; 
    ''' wenn inCludingSumRoles = true (default : false) , dann werden auch die Summary Roles ausgegeben
    ''' </summary>
    ''' <param name="includingSumRoles"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoleNames(Optional ByVal includingSumRoles As Boolean = False) As Collection

        Get
            Dim phase As clsPhase
            Dim aufbauRollen As New Collection
            Dim summaryRoles As Collection
            Dim roleName As String
            Dim hrole As clsRolle
            Dim p As Integer, r As Integer

            'Dim ende As Integer


            If Me._Dauer > 0 Then

                For p = 0 To AllPhases.Count - 1
                    phase = AllPhases.Item(p)
                    With phase
                        For r = 1 To .countRoles
                            hrole = .getRole(r)
                            If hrole.summe > 0 Then
                                roleName = hrole.name

                                '
                                ' das ist performanter als der Weg über try .. catch 
                                '
                                If Not aufbauRollen.Contains(roleName) Then
                                    aufbauRollen.Add(roleName, roleName)
                                End If

                                If includingSumRoles Then
                                    summaryRoles = RoleDefinitions.getSummaryRoles(roleName)
                                    For Each summaryRole As String In summaryRoles
                                        If Not aufbauRollen.Contains(summaryRole) Then
                                            aufbauRollen.Add(summaryRole, summaryRole)
                                        End If
                                    Next
                                End If

                            End If
                        Next r
                    End With
                Next p

            End If


            getRoleNames = aufbauRollen

        End Get

    End Property


    '
    ''' <summary>
    ''' gibt für Phase 1 ... n die Werte startoffset, dauer zurück 
    ''' Array hat die Dimension 2*anzPhasen -1 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseInfos() As Double()

        Get
            Dim anzPhasen As Integer
            Dim cphase As clsPhase
            Dim tmpvalues() As Double

            anzPhasen = AllPhases.Count
            ReDim tmpvalues(2 * anzPhasen - 1)

            For p = 0 To anzPhasen - 1

                cphase = AllPhases.Item(p)
                tmpvalues(p * 2) = cphase.startOffsetinDays
                tmpvalues(p * 2 + 1) = cphase.dauerInDays

            Next

            getPhaseInfos = tmpvalues

        End Get

    End Property

    Public ReadOnly Property getMilestoneColors() As Double()
        Get
            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein
            Dim tmpvalues() As Double
            Dim colorIndex As Integer
            Dim anzahlMilestones As Integer = 0

            For p = 1 To Me.CountPhases
                anzahlMilestones = anzahlMilestones + Me.getPhase(p).countMilestones
            Next

            If anzahlMilestones > 0 Then

                ReDim tmpvalues(anzahlMilestones - 1)

                Dim index As Integer = 0
                For p = 1 To Me.CountPhases
                    cphase = Me.getPhase(p)

                    For r = 1 To cphase.countMilestones
                        cresult = cphase.getMilestone(r)
                        colorIndex = cresult.getBewertung(1).colorIndex
                        tmpvalues(index) = colorIndex
                        index = index + 1
                    Next r

                Next p

            Else
                ReDim tmpvalues(0)
                tmpvalues(0) = 0
            End If

            getMilestoneColors = tmpvalues

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste an Meilenstein-Namen zurück (elem-Name, ohne Breadcrumb ...) 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneNames As Collection
        Get
            Dim hry As clsHierarchy = Me.hierarchy
            Dim tmpCollection As New Collection
            Dim firstMilestone As Integer = hry.getIndexOf1stMilestone


            If firstMilestone > 0 Then
                For ix As Integer = firstMilestone To hry.count
                    Dim msName As String = hry.nodeItem(ix).elemName
                    If Not tmpCollection.Contains(msName) Then
                        tmpCollection.Add(msName, msName)
                    End If
                Next
            End If


            getMilestoneNames = tmpCollection
        End Get
    End Property

    ''' <summary>
    ''' liefert eine Liste der vorkommenden Meilenstein Kategorien im Projekt zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneCategoryNames As Collection
        Get
            Dim hry As clsHierarchy = Me.hierarchy
            Dim tmpCollection As New Collection
            Dim firstMilestone As Integer = hry.getIndexOf1stMilestone


            If firstMilestone > 0 Then
                For ix As Integer = firstMilestone To hry.count
                    Dim tmpID As String = hry.getIDAtIndex(ix)
                    Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(tmpID)
                    If Not IsNothing(cMilestone) Then
                        Dim catName As String = cMilestone.appearance
                        If Not tmpCollection.Contains(catName) Then
                            tmpCollection.Add(catName, catName)
                        End If
                    End If

                Next
            End If

            getMilestoneCategoryNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste an ElemIDs von Meilensteinen zurück, die zu der angegebenen Category / appearance gehören  
    ''' </summary>
    ''' <param name="category"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneIDsWithCat(ByVal category As String) As Collection
        Get
            Dim hry As clsHierarchy = Me.hierarchy
            Dim tmpCollection As New Collection
            Dim firstMilestone As Integer = hry.getIndexOf1stMilestone

            If firstMilestone > 0 Then
                For ix As Integer = firstMilestone To hry.count
                    Dim msID As String = hry.getIDAtIndex(ix)
                    Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(msID)

                    If Not IsNothing(cMilestone) Then

                        If cMilestone.appearance = category Then

                            If Not tmpCollection.Contains(msID) Then
                                tmpCollection.Add(msID, msID)
                            End If

                        End If
                    End If

                Next
            End If

            getMilestoneIDsWithCat = tmpCollection
        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn das Projekt mindestens einen MEilenstein der angegebenen Category enthält 
    ''' </summary>
    ''' <param name="category"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsMilestoneCategory(ByVal category As String) As Boolean
        Get

            Dim hry As clsHierarchy = Me.hierarchy
            Dim found As Boolean = False
            Dim firstMilestone As Integer = hry.getIndexOf1stMilestone
            Dim ix As Integer = firstMilestone

            If firstMilestone > 0 Then
                While ix <= hry.count And Not found
                    Dim msID As String = hry.getIDAtIndex(ix)
                    Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(msID)

                    If Not IsNothing(cMilestone) Then
                        If cMilestone.appearance = category Then
                            found = True
                        End If
                    End If

                    ix = ix + 1
                End While

            End If

            containsMilestoneCategory = found

        End Get
    End Property

    ''' <summary>
    ''' liefert eine Liste der vorkommenden Meilenstein Kategorien im Projekt zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseCategoryNames As Collection
        Get
            Dim hry As clsHierarchy = Me.hierarchy
            Dim tmpCollection As New Collection
            Dim lastPhaseIx As Integer = hry.getIndexOf1stMilestone - 1
            If lastPhaseIx < 0 Then
                ' es gibt keine Meilensteine 
                lastPhaseIx = hry.count
            End If


            For ix As Integer = 1 To lastPhaseIx
                Dim tmpID As String = hry.getIDAtIndex(ix)
                Dim cPhase As clsPhase = Me.getPhaseByID(tmpID)
                If Not IsNothing(cPhase) Then
                    Dim catName As String = cPhase.appearance
                    If Not tmpCollection.Contains(catName) Then
                        tmpCollection.Add(catName, catName)
                    End If
                End If

            Next


            getPhaseCategoryNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste an ElemIDs von Phasen zurück, die zu der angegebenen Category / appearance gehören  
    ''' </summary>
    ''' <param name="category"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseIDsWithCat(ByVal category As String) As Collection
        Get
            Dim hry As clsHierarchy = Me.hierarchy
            Dim tmpCollection As New Collection
            Dim lastPhaseIx As Integer = hry.getIndexOf1stMilestone - 1
            If lastPhaseIx < 0 Then
                lastPhaseIx = hry.count
            End If

            For ix As Integer = 1 To lastPhaseIx
                Dim phID As String = hry.getIDAtIndex(ix)
                Dim cPhase As clsPhase = Me.getPhaseByID(phID)

                If Not IsNothing(cPhase) Then

                    If cPhase.appearance = category Then

                        If Not tmpCollection.Contains(phID) Then
                            tmpCollection.Add(phID, phID)
                        End If

                    End If
                End If

            Next

            getPhaseIDsWithCat = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn das Projekt mindesten einen Phase der angegebenen Category enthält 
    ''' </summary>
    ''' <param name="category"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsPhaseCategory(ByVal category As String) As Boolean
        Get

            Dim hry As clsHierarchy = Me.hierarchy
            Dim found As Boolean = False
            Dim lastPhaseIx As Integer = hry.getIndexOf1stMilestone - 1
            If lastPhaseIx < 0 Then
                lastPhaseIx = hry.count
            End If
            Dim ix As Integer = 1


            While ix <= lastPhaseIx And Not found
                Dim phID As String = hry.getIDAtIndex(ix)
                Dim cphase As clsPhase = Me.getPhaseByID(phID)

                If Not IsNothing(cphase) Then
                    If cphase.appearance = category Then
                        found = True
                    End If
                End If

                ix = ix + 1
            End While



            containsPhaseCategory = found

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste aller im Projekt vorkommenden Phasen Namen zurück ..
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseNames As Collection
        Get

            Dim tmpCollection As New Collection
            Dim lastPhase As Integer = Me.hierarchy.getIndexOf1stMilestone - 1
            If lastPhase < 0 Then
                lastPhase = Me.hierarchy.count
            End If


            If lastPhase > 0 Then
                For ix As Integer = 1 To lastPhase
                    Dim phName As String = Me.hierarchy.nodeItem(ix).elemName

                    If Not tmpCollection.Contains(phName) And phName <> elemNameOfElemID(rootPhaseName) Then
                        tmpCollection.Add(phName, phName)
                    End If
                Next
            End If

            getPhaseNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt zum betreffenden Projekt eine nach dem Datum aufsteigend sortierte Liste der Meilensteine zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns>nach Datum sortierte Liste der MEilensteine im Projekt </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestones As SortedList(Of Date, String)
        Get
            Dim tmpValues As New SortedList(Of Date, String)
            Dim tmpDate As Date
            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein

            For p = 1 To Me.CountPhases
                cphase = Me.getPhase(p)

                For r = 1 To cphase.countMilestones
                    cresult = cphase.getMilestone(r)
                    tmpDate = cresult.getDate

                    Dim ok As Boolean = False
                    Do While tmpValues.ContainsKey(tmpDate)
                        tmpDate = tmpDate.AddMilliseconds(1)
                    Loop
                    ' jetzt gibt es tmpdate noch nicht in der Liste ...
                    tmpValues.Add(tmpDate, cresult.nameID)
                Next r

            Next p

            getMilestones = tmpValues

        End Get
    End Property

    ''' <summary>
    ''' gibt zum betreffenden Projekt eine nach dem Offset aufsteigend sortierte Liste der Meilensteine zurück 
    ''' wird benötigt, wo ein relativer Vergleich der MEilensteine erforderlich ist 
    ''' bei Gleichheit wird ein Koorktur Faktor kleiner 1 addiert, so dass es immer eindeutige Werte gibt  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneOffsets As SortedList(Of Double, String)
        Get
            Dim tmpValues As New SortedList(Of Double, String)
            Dim tmpOffset As Double
            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein

            For p = 1 To Me.CountPhases
                cphase = Me.getPhase(p)

                For r = 1 To cphase.countMilestones
                    cresult = cphase.getMilestone(r)
                    tmpOffset = cphase.startOffsetinDays + cresult.offset

                    Dim korrFaktor As Double = 0.5
                    Do While tmpValues.ContainsKey(tmpOffset)
                        tmpOffset = tmpOffset + korrFaktor
                        korrFaktor = (1 - korrFaktor) * 0.5
                    Loop
                    ' jetzt gibt es tmpOffset noch nicht in der Liste ...
                    tmpValues.Add(tmpOffset, cresult.nameID)
                Next r

            Next p

            getMilestoneOffsets = tmpValues

        End Get
    End Property

    ''' <summary>
    ''' gibt eine sortierte Liste an Deliverables zurück; 
    ''' sortier-Kriterium ist der Name des Deliverables, value ist die nameID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDeliverables As SortedList(Of String, String)
        Get
            Dim tmpValues As New SortedList(Of String, String)
            Dim tmpDeliverable As String
            Dim tmpNameID As String
            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein

            For p = 1 To Me.CountPhases
                cphase = Me.getPhase(p)

                For r = 1 To cphase.countMilestones
                    cresult = cphase.getMilestone(r)
                    tmpNameID = cresult.nameID

                    For d = 1 To cresult.countDeliverables
                        tmpDeliverable = cresult.getDeliverable(d)
                        If tmpValues.ContainsKey(tmpDeliverable) Then
                            tmpDeliverable = tmpDeliverable & "(" & tmpNameID & ")"
                        Else
                            ' nichts tun, Deliverable existiert noch nicht ...
                        End If

                        ' jetzt ist sichergestellt, dass das Deliverable noch nicht existiert 
                        ' bzw. das Deliverable dieser NameID bereits aufgenommen ist 
                        If Not tmpValues.ContainsKey(tmpDeliverable) Then
                            tmpValues.Add(tmpDeliverable, tmpNameID)
                        Else
                            ' nichts mehr tun - Deliverable dieser NameID ist schon drin 

                        End If
                    Next d

                Next r

            Next p

            getDeliverables = tmpValues

        End Get
    End Property

    '
    ' übergibt in getPersonalKosten die Personal Kosten der Rolle <roleid> über den Projektzeitraum
    '
    Public ReadOnly Property getPersonalKosten(roleID As Object, Optional ByVal inclSubRoles As Boolean = False) As Double()
        Get
            Dim costValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer

            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim lookforIndex As Boolean
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim tagessatz As Double
            Dim faktor As Double = 1
            Dim dimension As Integer
            Dim roleUID As Integer
            Dim roleName As String = ""

            Dim roleIDs As New SortedList(Of Integer, Double)

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If


            If _Dauer > 0 Then
                lookforIndex = IsNumeric(roleID)

                If IsNumeric(roleID) Then
                    roleUID = CInt(roleID)
                    roleName = RoleDefinitions.getRoledef(roleUID).name
                Else
                    If RoleDefinitions.containsName(CStr(roleID)) Then
                        roleUID = RoleDefinitions.getRoledef(CStr(roleID)).UID
                        roleName = CStr(roleID)
                    End If
                End If

                ' jetzt prüfen, ob es inkl aller SubRoles sein soll 
                If inclSubRoles Then
                    roleIDs = RoleDefinitions.getSubRoleIDsOf(roleName, type:=PTcbr.all)
                Else
                    roleIDs.Add(roleUID, 1.0)
                End If

                ReDim costValues(_Dauer - 1)


                For Each srkvp As KeyValuePair(Of Integer, Double) In roleIDs
                    roleName = RoleDefinitions.getRoledef(srkvp.Key).name

                    Dim listOfPhases As Collection = Me.rcLists.getPhasesWithRole(roleName, False)
                    anzPhasen = listOfPhases.Count

                    For p = 1 To anzPhasen
                        phase = Me.getPhaseByID(CStr(listOfPhases.Item(p)))

                        With phase
                            ' Off1
                            anzRollen = .countRoles
                            phasenStart = .relStart - 1
                            'phasenEnde = .relEnde - 1


                            For r = 1 To anzRollen
                                role = .getRole(r)

                                With role

                                    If .RollenTyp = srkvp.Key Then

                                        tagessatz = .tagessatzIntern
                                        dimension = .getDimension
                                        ReDim tempArray(dimension)
                                        tempArray = .Xwerte

                                        For i = phasenStart To phasenStart + dimension
                                            costValues(i) = costValues(i) + tempArray(i - phasenStart) * tagessatz * faktor / 1000
                                        Next i

                                    End If

                                End With ' role

                            Next r

                        End With ' phase

                    Next
                Next

            Else
                ReDim costValues(0)
                costValues(0) = 0
            End If

            getPersonalKosten = costValues

        End Get
    End Property

    ''' <summary>
    ''' gibt den anteiligen Wert der Rolle/Kostenart in der betreffenden Phase an den Gesamtkosten zurück;
    ''' kann verwendet werden, um Best Practice Projekte on-the-fly zu definieren  
    ''' </summary>
    ''' <param name="phaseID"></param>
    ''' <param name="rcName"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPercentShareOFTotalCost(ByVal phaseID As String, ByVal rcName As String, ByVal type As Integer) As Double
        Get
            Dim tmpResult As Double = 0.0
            Dim totalCost As Double = Me.getSummeKosten()
            Dim cphase As clsPhase = Me.getPhaseByID(phaseID)
            Dim role As clsRolle
            Dim cost As clsKostenart
            Dim teilWert As Double

            If totalCost > 0 And Not IsNothing(cphase) Then
                If type = ptElementTypen.roles Then
                    role = cphase.getRole(rcName)
                    If Not IsNothing(role) Then
                        teilWert = role.Xwerte.Sum * role.tagessatzIntern
                    Else
                        teilWert = 0
                    End If
                    tmpResult = teilWert / totalCost

                ElseIf type = ptElementTypen.costs Then
                    cost = cphase.getCost(rcName)
                    If Not IsNothing(cost) Then
                        teilWert = cost.Xwerte.Sum
                    Else
                        teilWert = 0
                    End If
                    tmpResult = teilWert / totalCost
                End If

            End If

            getPercentShareOFTotalCost = tmpResult

        End Get
    End Property

    '
    ' übergibt in KostenBedarf die Werte der Kostenart <costId>
    '
    Public ReadOnly Property getKostenBedarf(CostID As Object) As Double()

        Get
            Dim costValues() As Double
            Dim anzKostenarten As Integer
            Dim anzPhasen As Integer
            Dim found As Boolean
            Dim i As Integer, p As Integer, k As Integer
            Dim phase As clsPhase
            Dim cost As clsKostenart
            Dim lookforIndex As Boolean, isPersCost As Boolean
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim dimension As Integer


            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)

                lookforIndex = IsNumeric(CostID)
                isPersCost = False

                If lookforIndex Then
                    If CostID = CostDefinitions.Count Then
                        isPersCost = True
                    End If
                Else
                    If CostID = "Personalkosten" Then
                        isPersCost = True
                    End If
                End If

                If isPersCost Then
                    ' costvalues = AllPersonalKosten
                    costValues = Me.getAllPersonalKosten
                Else

                    anzPhasen = AllPhases.Count

                    For p = 0 To anzPhasen - 1
                        phase = AllPhases.Item(p)
                        With phase
                            ' Off1
                            anzKostenarten = .countCosts
                            phasenStart = .relStart - 1
                            'phasenEnde = .relEnde - 1


                            For k = 1 To anzKostenarten
                                cost = .getCost(k)
                                found = False

                                With cost
                                    If lookforIndex Then
                                        If .KostenTyp = CostID Then
                                            found = True
                                        End If
                                    Else
                                        If .name = CostID Then
                                            found = True
                                        End If
                                    End If
                                    If found Then
                                        dimension = .getDimension
                                        ReDim tempArray(dimension)
                                        tempArray = .Xwerte
                                        For i = phasenStart To phasenStart + dimension

                                            costValues(i) = costValues(i) + tempArray(i - phasenStart)


                                        Next i
                                    End If
                                End With ' cost

                            Next k

                        End With ' phase

                    Next p ' Loop über alle Phasen
                End If
            Else
                ReDim costValues(0)
                costValues(0) = 0
            End If

            getKostenBedarf = costValues


        End Get

    End Property

    '
    ' übergibt in KostenBedarf die Werte der Kostenart <costId>
    '
    Public ReadOnly Property getKostenBedarfNew(CostID As Object) As Double()

        Get
            Dim costValues() As Double
            Dim anzKostenarten As Integer
            Dim anzPhasen As Integer
            Dim found As Boolean
            Dim i As Integer, p As Integer, k As Integer
            Dim phase As clsPhase
            Dim cost As clsKostenart
            Dim lookforIndex As Boolean, isPersCost As Boolean
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim dimension As Integer
            Dim costUID As Integer = 0
            Dim costName As String = ""


            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)

                lookforIndex = IsNumeric(CostID)
                isPersCost = False

                If lookforIndex Then
                    costUID = CInt(CostID)
                    costName = CostDefinitions.getCostdef(costUID).name
                    If CostID = CostDefinitions.Count Then
                        isPersCost = True
                    End If
                Else
                    If CostDefinitions.containsName(CStr(CostID)) Then
                        costUID = CostDefinitions.getCostdef(CStr(CostID)).UID
                        costName = CStr(CostID)
                    End If
                    If CostID = "Personalkosten" Then
                        isPersCost = True
                    End If
                End If

                If isPersCost Then
                    ' costvalues = AllPersonalKosten
                    costValues = Me.getAllPersonalKosten
                Else
                    Dim listOfPhases As Collection = Me.rcLists.getPhasesWithCost(costName)
                    anzPhasen = listOfPhases.Count
                    
                    For p = 1 To anzPhasen
                        phase = Me.getPhaseByID(CStr(listOfPhases.Item(p)))

                        With phase
                            ' Off1
                            anzKostenarten = .countCosts
                            phasenStart = .relStart - 1
                            'phasenEnde = .relEnde - 1


                            For k = 1 To anzKostenarten
                                cost = .getCost(k)
                                found = False

                                With cost

                                    If .KostenTyp = costUID Then

                                        dimension = .getDimension
                                        ReDim tempArray(dimension)
                                        tempArray = .Xwerte

                                        For i = phasenStart To phasenStart + dimension
                                            costValues(i) = costValues(i) + tempArray(i - phasenStart)
                                        Next i

                                    End If

                                End With ' cost

                            Next k

                        End With ' phase

                    Next p ' Loop über alle Phasen
                End If
            Else
                ReDim costValues(0)
                costValues(0) = 0
            End If

            getKostenBedarfNew = costValues


        End Get

    End Property

    '
    ' übergibt in getUsedKosten eine Collection von Kostenarten Definitionen, 
    ' das sind alle Kostenarten, die in den Phasen vorkommen und einen Bedarf von größer Null haben
    '
    Public ReadOnly Property getCostNames() As Collection

        Get
            Dim phase As clsPhase
            Dim aufbauKosten As New Collection
            Dim costname As String
            Dim hcost As clsKostenart
            Dim p As Integer, k As Integer

            'Dim ende As Integer

            If _Dauer > 0 Then
                For p = 0 To AllPhases.Count - 1
                    phase = AllPhases.Item(p)
                    With phase
                        For k = 1 To .countCosts
                            hcost = .getCost(k)
                            If hcost.summe > 0 Then
                                costname = hcost.name
                                '
                                ' das ist performanter als über try .. catch 
                                '
                                If Not aufbauKosten.Contains(costname) Then
                                    aufbauKosten.Add(costname, costname)
                                End If
                                'Try
                                '    aufbauKosten.Add(costname, costname)
                                'Catch ex As Exception

                                'End Try

                            End If
                        Next k
                    End With
                Next p

            End If


            getCostNames = aufbauKosten

        End Get

    End Property


    ''' <summary>
    ''' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSummeKosten() As Double

        Get
            Dim costValues() As Double
            Dim ErgebnisListe As New Collection
            Dim costSum As Double
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String

            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)
                costValues = Me.getAllPersonalKosten

                costSum = 0
                For i = 0 To _Dauer - 1
                    costSum = costSum + costValues(i)
                    costValues(i) = 0
                Next i
                '
                ' jetzt sind in der Summe die Personalkosten drin ....
                '

                ' Jetzt werden die einzelnen Kostenarten auf die gleiche Art und Weise geholt
                ErgebnisListe = Me.getCostNames

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = CStr(ErgebnisListe.Item(r))
                    costValues = Me.getKostenBedarf(costname)
                    For i = 0 To _Dauer - 1
                        costSum = costSum + costValues(i)
                        costValues(i) = 0
                    Next i
                Next r

                getSummeKosten = costSum

            Else
                getSummeKosten = 0
            End If

        End Get

    End Property


    ''' <summary>
    ''' berechnet die Summe nur bis zum index.ten Monaten 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSummeKosten(ByVal index As Integer) As Double

        Get
            Dim costValues() As Double
            Dim ErgebnisListe As New Collection
            Dim costSum As Double
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String

            If _Dauer > 0 Then

                If index > _Dauer - 1 Then
                    index = _Dauer - 1
                End If

                ReDim costValues(_Dauer - 1)
                costValues = Me.getAllPersonalKosten

                costSum = 0
                For i = 0 To index

                    costSum = costSum + costValues(i)

                Next i
                '
                ' jetzt sind in der Summe die Personalkosten drin ....
                '

                ' Jetzt werden die einzelnen Kostenarten auf die gleiche Art und Weise geholt
                ErgebnisListe = Me.getCostNames

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = ErgebnisListe.Item(r).ToString

                    ReDim costValues(_Dauer - 1)
                    costValues = Me.getKostenBedarf(costname)
                    For i = 0 To index

                        costSum = costSum + costValues(i)

                    Next i
                Next r

                getSummeKosten = costSum

            Else
                getSummeKosten = 0
            End If

        End Get

    End Property

    '
    ' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    '
    Public ReadOnly Property getGesamtKostenBedarf() As Double()

        Get
            Dim costValues() As Double, tmpValues() As Double
            Dim ErgebnisListe As New Collection
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String


            ReDim costValues(_Dauer - 1)
            ReDim tmpValues(_Dauer - 1)

            If _Dauer > 0 Then

                costValues = Me.getAllPersonalKosten
                '
                ' jetzt sind in costValues die Personalkosten drin ....
                '

                ' Jetzt werden die einzelnen Kostenarten auf die gleiche Art und Weise geholt
                ErgebnisListe = Me.getCostNames

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = CStr(ErgebnisListe.Item(r))
                    tmpValues = Me.getKostenBedarf(costname)
                    For i = 0 To _Dauer - 1
                        costValues(i) = costValues(i) + tmpValues(i)
                        tmpValues(i) = 0
                    Next i
                Next r

            End If

            getGesamtKostenBedarf = costValues

        End Get

    End Property

    '
    ' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    '
    Public ReadOnly Property getGesamtAndereKosten() As Double()

        Get
            Dim costValues() As Double, tmpValues() As Double
            Dim ErgebnisListe As New Collection
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String


            ReDim costValues(_Dauer - 1)
            ReDim tmpValues(_Dauer - 1)

            If _Dauer > 0 Then

                ' Jetzt werden die einzelnen Kostenarten geholt
                ErgebnisListe = Me.getCostNames

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = CStr(ErgebnisListe.Item(r))
                    tmpValues = Me.getKostenBedarf(costname)
                    For i = 0 To _Dauer - 1
                        costValues(i) = costValues(i) + tmpValues(i)
                        tmpValues(i) = 0
                    Next i
                Next r

            End If

            getGesamtAndereKosten = costValues

        End Get

    End Property

    '
    ' übergibt in getSummeRessourcen den Ressourcen Bedarf in Mann-Monaten  die Werte der Kostenart <roleId>
    '
    Public ReadOnly Property getSummeRessourcen() As Double

        Get
            Dim roleValues() As Double
            Dim ErgebnisListe As New Collection
            Dim roleSum As Double
            Dim anzRollen As Integer
            Dim i As Integer, r As Integer
            Dim roleName As String


            If _Dauer > 0 Then

                ReDim roleValues(_Dauer - 1)

                roleSum = 0

                ' Jetzt werden die einzelnen Rollen aufsummiert
                ErgebnisListe = Me.getRoleNames
                anzRollen = ErgebnisListe.Count

                For r = 1 To anzRollen
                    roleName = CStr(ErgebnisListe.Item(r))
                    roleValues = Me.getRessourcenBedarf(roleName)
                    For i = 0 To _Dauer - 1
                        roleSum = roleSum + roleValues(i)
                        roleValues(i) = 0
                    Next i
                Next r

                getSummeRessourcen = roleSum

            Else
                getSummeRessourcen = 0
            End If

        End Get

    End Property


    ''' <summary>
    ''' liefert einen Array der Länge dauer-1 mit den monatlichen Gesamt Ressourcenbedarfen 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getAlleRessourcen() As Double()

        Get
            Dim roleValues() As Double
            Dim alleValues() As Double
            Dim ErgebnisListe As New Collection
            Dim anzRollen As Integer
            Dim i As Integer, r As Integer
            Dim roleName As String


            If _Dauer > 0 Then

                ReDim roleValues(_Dauer - 1)
                ReDim alleValues(_Dauer - 1)


                ' Jetzt werden die einzelnen Rollen aufsummiert
                ErgebnisListe = Me.getRoleNames
                anzRollen = ErgebnisListe.Count

                For r = 1 To anzRollen
                    roleName = CStr(ErgebnisListe.Item(r))
                    roleValues = Me.getRessourcenBedarf(roleName)
                    For i = 0 To _Dauer - 1
                        alleValues(i) = alleValues(i) + roleValues(i)
                        roleValues(i) = 0
                    Next i
                Next r

                getAlleRessourcen = alleValues

            Else
                ReDim alleValues(0)
                getAlleRessourcen = alleValues
            End If

        End Get

    End Property



    ''' <summary>
    ''' gibt die Personalkosten des betreffenden Projektes zurück ; zugrundgelegt wird der interne Tagessatz 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllPersonalKosten() As Double()

        Get
            Dim costValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim tagessatz As Double
            Dim faktor As Double = 1
            Dim dimension As Integer

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Or awinSettings.kapaEinheit = "PD" Then
                faktor = 1
            Else
                faktor = 1
            End If


            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)


                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzRollen = .countRoles
                        phasenStart = .relStart - 1
                        'phasenEnde = .relEnde - 1


                        For r = 1 To anzRollen
                            role = .getRole(r)

                            With role
                                tagessatz = .tagessatzIntern
                                dimension = .getDimension
                                ReDim tempArray(dimension)
                                tempArray = .Xwerte
                                For i = phasenStart To phasenStart + dimension
                                    costValues(i) = costValues(i) + tempArray(i - phasenStart) * tagessatz * faktor / 1000
                                Next i

                            End With ' role

                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen



            Else

                ReDim costValues(0)
                costValues(0) = 0

            End If

            getAllPersonalKosten = costValues

        End Get

    End Property

    Public Overridable Property earliestStart() As Integer

        Get
            earliestStart = _earliestStart
        End Get
        Set(value As Integer)
            If value > 0 Then
                Throw New ArgumentException("Earliest Start kann nicht nach dem Startdatum liegen")
            Else
                _earliestStart = value
            End If

        End Set

    End Property


    Public Overridable Property latestStart() As Integer

        Get
            latestStart = _latestStart
        End Get
        Set(value As Integer)

            If value < 0 Then
                Throw New ArgumentException("latest Start kann nicht vor dem Startdatum liegen")
            Else
                _latestStart = value
            End If

        End Set

    End Property



    Public Sub New()

        AllPhases = New List(Of clsPhase)
        ' Änderung tk 31.3.15
        hierarchy = New clsHierarchy

        ' Änderung / Ergänzung tk 20.09.16
        rcLists = New clsListOfCostAndRoles

        relStart = 1
        _Dauer = 0
        '_StartOffset = 0
        '_Start = 1
        _earliestStart = 0
        _latestStart = 0
        '_Status = ProjektStatus(0)
        Schrift = 12
        Schriftfarbe = RGB(0, 0, 0)

        ' die CustomFields initialisieren 
        _customDblFields = New SortedList(Of Integer, Double)
        _customStringFields = New SortedList(Of Integer, String)
        _customBoolFields = New SortedList(Of Integer, Boolean)


    End Sub



End Class
