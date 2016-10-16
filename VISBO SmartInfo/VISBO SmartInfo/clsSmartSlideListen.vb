Public Class clsSmartSlideListen

    ' um zu verhindern, dass der Speicherbedarf wegen sortierter String Listen sehr groß wird, 
    ' wird eine Hilfsliste eingeführt, die für jeden auftretenden Shape-Namen (eindeutig !) eine eindeutige lfdNr zuweist 
    Private planShapeIDs As SortedList(Of String, Integer)
    Private IDplanShapes As SortedList(Of Integer, String)
    Private cNList As SortedList(Of String, SortedList(Of Integer, Boolean))
    Private oNList As SortedList(Of String, SortedList(Of Integer, Boolean))
    Private sNList As SortedList(Of String, SortedList(Of Integer, Boolean))
    Private bCList As SortedList(Of String, SortedList(Of Integer, Boolean))
    Private aCList As SortedList(Of Integer, SortedList(Of Integer, Boolean))

    Public ReadOnly Property getUID(ByVal shapeName As String) As Integer
        Get
            Dim uid As Integer
            If planShapeIDs.ContainsKey(shapeName) Then
                uid = planShapeIDs.Item(shapeName)
            Else
                uid = planShapeIDs.Count + 1
                planShapeIDs.Add(shapeName, uid)
                IDplanShapes.Add(uid, shapeName)
            End If

            getUID = uid

        End Get
    End Property

    ''' <summary>
    ''' gibt den ShapeName zurück, der zur UID gehört; 
    ''' </summary>
    ''' <param name="UID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property getShapeNameOfUid(ByVal uid As Integer) As String
        Get
            Dim tmpStr As String = ""
            Dim tmpStrTest As String = ""
            Dim found As Boolean = False
            Dim index As Integer = 0

            tmpStr = IDplanShapes.Item(uid)

            '' für Testzwecke 
            'Do While index <= planShapeIDs.Count - 1 And Not found
            '    If planShapeIDs.ElementAt(index).Value = uid Then
            '        found = True
            '        tmpStrTest = planShapeIDs.ElementAt(index).Key
            '    Else
            '        index = index + 1
            '    End If

            'Loop

            'If tmpStr <> tmpStrTest Then
            '    Dim a As Integer = 0
            'End If

            getShapeNameOfUid = tmpStr

        End Get
    End Property
    ''' <summary>
    ''' fügt der Liste an classified Names einen weiteren Namen hinzu
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="cName"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addCN(ByVal cName As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If cNList.ContainsKey(cName) Then
            listOfShapeNames = cNList.Item(cName)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            cNList.Add(cName, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an original Names einen weiteren Namen hinzu
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="oName">original Name</param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addON(ByVal oName As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If oNList.ContainsKey(oName) Then
            listOfShapeNames = oNList.Item(oName)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            oNList.Add(oName, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an Short Names einen weiteren Namen hinzu
    ''' wenn der leer ist, wird stattdessen die uid genommen 
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben
    ''' </summary>
    ''' <param name="sName"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addSN(ByVal sName As String, shapeName As String)


        Dim uid As Integer = Me.getUID(shapeName)
        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If IsNothing(sName) Then
            sName = uid.ToString
        ElseIf sName.Trim.Length = 0 Then
            sName = uid.ToString
        End If

        If sNList.ContainsKey(sName) Then
            listOfShapeNames = sNList.Item(sName)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            sNList.Add(sName, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an BreadCrumbs Names einen weiteren bc hinzu
    ''' wenn der schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="bCrumb"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addBC(ByVal bCrumb As String, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim fullbCrumb As String = "(" & getPnameFromShpName(shapeName) & ")" & _
            bCrumb.Replace("#", " - ") & getElemNameFromShpName(shapeName)


        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        If bCList.ContainsKey(fullbCrumb) Then
            listOfShapeNames = bCList.Item(fullbCrumb)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            bCList.Add(fullbCrumb, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' fügt der Liste an Ampelfarben eine weitere (0,1,2,3) hinzu
    ''' wenn die schon existiert, wird die Liste an shapeNames ergänzt; statt ShapeName wird dessen uid geschrieben  
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Public Sub addAC(ByVal ampelColor As Integer, shapeName As String)

        Dim uid As Integer = Me.getUID(shapeName)

        Dim listOfShapeNames As SortedList(Of Integer, Boolean)

        ' konsistent machen ... wenn die Farbe nicht erkannt werden kann, wird sie wie <nicht gesetzt> behandelt 
        If ampelColor < 0 Or ampelColor > 3 Then
            ampelColor = 0
        End If

        If aCList.ContainsKey(ampelColor) Then
            listOfShapeNames = aCList.Item(ampelColor)
            If listOfShapeNames.ContainsKey(uid) Then
                ' nichts tun , ist schon drin ...
            Else
                ' aufnehmen ; der bool'sche Value hat aktuell keine Bedeutung 
                listOfShapeNames.Add(uid, True)
            End If
        Else
            ' dann muss das erste aufgenommen werden 
            listOfShapeNames = New SortedList(Of Integer, Boolean)
            listOfShapeNames.Add(uid, True)
            aCList.Add(ampelColor, listOfShapeNames)
        End If

    End Sub

    ''' <summary>
    ''' gibt für die angegebene Ampelfarbe die Namen alle Shapes zurück, die diese Ampelfarbe haben 
    ''' leere Collection, wenn es keine Shapes dieser Farbe gibt
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapeNamesWithColor(ByVal ampelColor As Integer) As Collection
        Get
            Dim tmpCollection As New Collection

            Try
                If Not IsNothing(aCList) Then
                    Dim uidsWithColor As SortedList(Of Integer, Boolean) = aCList.Item(ampelColor)

                    If Not IsNothing(uidsWithColor) Then
                        ' jetzt sind in der uidList alle ShapeUIDs aufgeführt - die müssen jetzt durch ihre ShapeNames ersetzt werden 
                        For Each kvp As KeyValuePair(Of Integer, Boolean) In uidsWithColor

                            Dim shpName As String = Me.getShapeNameOfUid(kvp.Key)

                            If shpName.Trim.Length > 0 Then
                                If Not tmpCollection.Contains(shpName) Then
                                    tmpCollection.Add(shpName, shpName)
                                End If
                            End If

                        Next
                    End If
                    
                End If
            Catch ex As Exception

            End Try
            

            getShapeNamesWithColor = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' bekommt als Input eine Menge von selektierten Namen , classified, Short, Original, etc. 
    ''' gibt als Output die korrespondierenden Shape-NAmen
    ''' Achtung: Anzahl Input Elemente muss nicht Anzahl Output Elemente sein;  
    ''' </summary>
    ''' <param name="nameArray"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapesNames(ByVal nameArray() As String, _
                                                ByVal type As Integer) As Collection
        Get
            Dim tmpCollection As New Collection

            Dim NList As SortedList(Of String, SortedList(Of Integer, Boolean))
            Dim alleUIDs As New SortedList(Of Integer, Boolean)
            Dim anzahlNames As Integer = nameArray.Length

            Select Case type
                Case pptInfoType.cName
                    NList = cNList
                Case pptInfoType.oName
                    NList = oNList
                Case pptInfoType.sName
                    NList = sNList
                Case pptInfoType.bCrumb
                    NList = bCList
                Case Else
                    NList = cNList
            End Select

            For i As Integer = 0 To anzahlNames - 1

                Dim uidList As SortedList(Of Integer, Boolean) = NList.Item(nameArray(i))

                For Each kvp As KeyValuePair(Of Integer, Boolean) In uidList

                    If Not alleUIDs.ContainsKey(kvp.Key) Then
                        alleUIDs.Add(kvp.Key, kvp.Value)
                    End If

                Next

            Next

            ' jetzt sind in der uidList alle ShapeUIDs aufgeführt - die müssen jetzt durch ihre ShapeNames ersetzt werden 
            For Each kvp As KeyValuePair(Of Integer, Boolean) In alleUIDs

                Dim shpName As String = Me.getShapeNameOfUid(kvp.Key)

                If shpName.Trim.Length > 0 Then
                    If Not tmpCollection.Contains(shpName) Then
                        tmpCollection.Add(shpName, shpName)
                    End If
                End If

            Next

            getShapesNames = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt eine Liste zurück an Element-Namen, die den Suchstr enthalten und ausserdem die übergebene Farben-Kennung haben
    ''' leere Liste, wenn es keine Entsprechung gibt  
    ''' </summary>
    ''' <param name="colorCode"></param>
    ''' <param name="suchStr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getNCollection(ByVal colorCode As Integer, _
                                         ByVal suchStr As String, _
                                         ByVal type As Integer) As Collection
        Get
            Dim NList As SortedList(Of String, SortedList(Of Integer, Boolean))

            Select Case type
                Case pptInfoType.cName
                    NList = cNList
                Case pptInfoType.oName
                    NList = oNList
                Case pptInfoType.sName
                    NList = sNList
                Case pptInfoType.bCrumb
                    NList = bCList
                Case Else
                    NList = cNList
            End Select

            Dim tmpCollection As New Collection
            Dim alleUIDsMitgesuchterFarbe As SortedList(Of Integer, Boolean)

            Dim txtRestriction As Boolean = False
            Dim colRestriction As Boolean = False

            ' gibt es eine Text Restriction, also muss der Name irgendwas enthalten ...
            If IsNothing(suchStr) Then
            ElseIf suchStr.Trim.Length = 0 Then
            Else
                txtRestriction = True
            End If

            ' gibt es eine Color Restriction, also sollen nur bestimmte Farben angezeigt werden 
            If colorCode < 1 Or colorCode > 15 Then
            Else
                colRestriction = True
            End If

            Dim uidList As SortedList(Of Integer, Boolean)

            ' erst wird die Liste an uids ermittelt, die den entsprechenden Farb-Code aufweisen
            ' dann wird untersucht, welche dieser uids ggf noch dem Suchstring entsprechen ... 

            alleUIDsMitgesuchterFarbe = New SortedList(Of Integer, Boolean)


            If colRestriction Then

                ' jetzt muss eine Schleife gemacht werden
                Dim singleFlag As Integer

                Do While colorCode > 0
                    If colorCode >= 8 Then
                        ' red Flag 
                        singleFlag = 3
                        If aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 8

                    ElseIf colorCode >= 4 Then
                        ' yellow flag 
                        singleFlag = 2
                        If aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 4

                    ElseIf colorCode >= 2 Then
                        ' green flag
                        singleFlag = 1
                        If aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 2

                    ElseIf colorCode >= 1 Then
                        ' nicht bewertet 
                        singleFlag = 0
                        If aCList.ContainsKey(singleFlag) Then
                            For Each kvp As KeyValuePair(Of Integer, Boolean) In aCList.Item(singleFlag)
                                alleUIDsMitgesuchterFarbe.Add(kvp.Key, kvp.Value)
                            Next
                        End If
                        colorCode = colorCode - 1
                    End If
                Loop


                If alleUIDsMitgesuchterFarbe.Count > 0 Then
                    ' es gibt Shapes - jetzt prüfen, ob es TextRestriktion gibt 

                    If txtRestriction Then
                        ' ermittle die UIDS, die den gesuchten Text enthalten , prüfe gleichzeitig, 
                        ' ob sie bereits in alleUIDSMitgesuchterFarbe sind ... 
                        ' trage die in ErgebnisListe ein 

                        ' Nlsit enthält die Namen, Original-NAmen, etc; jeweils mit einer Liste an UIDS, welche Elemente alle diesen 
                        ' einen Namen enthalten ; ggf kann aj z.B Montage mehrfach vorkommen - und die eine Montage UID hat die gesuchte Farbe, die andere nicht ... 
                        For Each listElem As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In NList

                            If listElem.Key.Contains(suchStr) Then
                                uidList = listElem.Value
                                For Each chkUID As KeyValuePair(Of Integer, Boolean) In uidList
                                    If alleUIDsMitgesuchterFarbe.ContainsKey(chkUID.Key) Then
                                        ' diese UID ist jetzt eine Ergebnis-UID , die sowhl die richtige Farbe als auch den richtigen Text-String hat 
                                        ' in listElem.key steht der gesuchte String .. 
                                        tmpCollection.Add(listElem.Key)
                                    End If
                                Next
                            End If

                        Next
                    Else
                        ' ermittle jetzt die Namen, Original-Namen für die Farb-UIDs
                        ' keine Text Restriktion
                        For Each listElem As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In NList

                            uidList = listElem.Value
                            For Each chkUID As KeyValuePair(Of Integer, Boolean) In uidList
                                If alleUIDsMitgesuchterFarbe.ContainsKey(chkUID.Key) Then
                                    ' diese UID ist jetzt eine Ergebnis-UID , die sowhl die richtige Farbe als auch den richtigen Text-String hat 
                                    ' in listElem.key steht der gesuchte String .. 
                                    tmpCollection.Add(listElem.Key)
                                End If
                            Next

                        Next
                    End If

                Else
                    ' nichts tun - alleUIDsMitgesuchterFarbe ist leer ...  
                End If

            Else
                ' keine Farb-Einschränkung - also einfach mal die cNList durchgehen 
                For Each listElem As KeyValuePair(Of String, SortedList(Of Integer, Boolean)) In NList

                    If txtRestriction Then
                        If listElem.Key.Contains(suchStr) Then
                            tmpCollection.Add(listElem.Key)
                        End If
                    Else
                        tmpCollection.Add(listElem.Key)
                    End If


                Next

            End If

            getNCollection = tmpCollection

        End Get
    End Property

    Public Sub New()
        planShapeIDs = New SortedList(Of String, Integer)
        IDplanShapes = New SortedList(Of Integer, String)
        cNList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        oNList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        sNList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        bCList = New SortedList(Of String, SortedList(Of Integer, Boolean))
        aCList = New SortedList(Of Integer, SortedList(Of Integer, Boolean))
    End Sub

End Class
