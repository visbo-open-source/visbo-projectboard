Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ProjectBoardBasic
Imports xlNS = Microsoft.Office.Interop.Excel
Public Class ucSearch



    Friend abkuerzung As String
    Friend showSearchListBox As Boolean = False

    Private dontFire As Boolean = False
    ' innerhalb der Klasse überall im Zugriff; Colorcode ist die Zahl , die sich ergibt , 
    ' wenn man die Werte 0, 1, 2, 3 als Potenzen von 2 und in Summe ausrechnet

    ' lokale Variable für die gewünschten Ampeln
    Private hShowTrafficLights(4) As Boolean

    ' wird in den entsprechenden Checkbox Routinen gesetzt 
    Private colorCode As Integer = 0

    ' wird im entsprechenden Suchfeld gesetzt 
    Private suchString As String = ""

    ' gibt an, ob bei der suche die gefundenen Elemente mit AMrker angezeigt werden sollen oder nicht .. 
    Friend showMarker As Boolean = False

    '' '' tk , 16.5.
    '' '' Private Const deltaAmpel As Integer = 50
    ' ''Private Const deltaAmpel As Integer = 0
    ' ''Private Const deltaSearchBox As Integer = 200
    ' ''Private Const smallHeight As Integer = 220

    '' '' steuert, wo der Text relatic zum Meilenstein , zur Phase platziert werden soll 
    '' '' MD: MilestoneDate, MT MilestoneText , PD PhaseDate, PT PhaseText
    ' ''Friend positionIndexMD As Integer = 5
    ' ''Friend positionIndexMT As Integer = 1
    ' ''Friend positionIndexPD As Integer = 8
    ' ''Friend positionIndexPT As Integer = 6

    ' hierin werden zu den angezeigten Namen in selListboxNames die shape-Namen gemerkt
    Friend shpNameSav As New SortedList(Of String, String)

    ''' <summary>
    ''' ist gecheckt, wenn ein Pfeil für das/die ausgewählten Elemente angezeigt werden soll
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CheckBxMarker_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBxMarker.CheckedChanged
        If CheckBxMarker.Checked Then
            ' alle selektierten Elemente jetzt mit Marker versehen
            showMarker = True
            Call createMarkerShapes(pptShapes:=selectedPlanShapes)
        Else
            showMarker = False
            Call deleteMarkerShapes()
        End If
    End Sub

    ''' <summary>
    ''' füllt die ListboxNames mit den Elementen, deren Ampel keine Bewertung hat
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub shwOhneLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwOhneLight.CheckedChanged

        Dim ampelColor As Integer = PTfarbe.none
        hshowTrafficLights(ampelColor) = shwOhneLight.Checked

        If shwOhneLight.Checked Then

        End If

        Call fülltListbox(hShowTrafficLights)

        Call faerbeShapes(ampelColor, shwOhneLight.Checked)
    End Sub

    ''' <summary>
    ''' füllt die listboxNames mit den Elementen, deren Ampel mit grün bewertet ist
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub shwGreenLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwGreenLight.CheckedChanged
        Dim ampelColor As Integer = PTfarbe.green
        hShowTrafficLights(ampelColor) = shwGreenLight.Checked

        If shwGreenLight.Checked Then

        End If

        Call fülltListbox(hShowTrafficLights)

        Call faerbeShapes(ampelColor, shwGreenLight.Checked)
    End Sub

    ''' <summary>
    '''  füllt die listboxNames mit den Elementen, deren Ampel mit gelb bewertet ist
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub shwYellowLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwYellowLight.CheckedChanged

        Dim ampelColor As Integer = PTfarbe.yellow
        hShowTrafficLights(ampelColor) = shwYellowLight.Checked

        If shwYellowLight.Checked Then

        End If

        Call fülltListbox(hShowTrafficLights)

        Call faerbeShapes(ampelColor, shwYellowLight.Checked)
    End Sub


    ''' <summary>
    '''  füllt die listboxNames mit den Elementen, deren Ampel mit rot bewertet ist
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub shwRedLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwRedLight.CheckedChanged
        Dim ampelColor As Integer = PTfarbe.red
        hShowTrafficLights(ampelColor) = shwRedLight.Checked

        If shwRedLight.Checked Then

        End If

        Call fülltListbox(hShowTrafficLights)

        Call faerbeShapes(ampelColor, shwRedLight.Checked)
    End Sub

    ''' <summary>
    ''' erstellt die Listbox aufgrund der Settings bei Ampeln, Radio-Button und Suchstr neu 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub fülltListbox(ByVal sTrafLights() As Boolean)

        If Not dontFire Then

            selListboxNames.Items.Clear()

            colorCode = calcColorCode(sTrafLights)

            Dim catCode As Integer


            ''
            '' hier muss die Textbox categoryList ausgelesen werden. Hier wird gefiltert
            ''
            If englishLanguage Then

                Select Case cathegoryList.SelectedItem
                    Case "Name"
                        catCode = pptInfoType.cName
                    Case "Responsibilities"
                        catCode = pptInfoType.responsible
                    Case "Original Name"
                        catCode = pptInfoType.oName
                    Case "Abbreviation"
                        catCode = pptInfoType.sName
                    Case "voller Name"
                        catCode = pptInfoType.bCrumb
                    Case "Deliverables"
                        catCode = pptInfoType.lUmfang
                    Case "manually Changed Dates"
                        catCode = pptInfoType.mvElement
                    Case "Resources"
                        catCode = pptInfoType.resources
                    Case "Cost"
                        catCode = pptInfoType.costs
                    Case "Overdue"
                        catCode = pptInfoType.overDue
                    Case Else
                        catCode = pptInfoType.cName
                End Select

            Else
                Select Case cathegoryList.SelectedItem
                    Case "Name"
                        catCode = pptInfoType.cName
                    Case "Verantwortlich"
                        catCode = pptInfoType.responsible
                    Case "Original Name"
                        catCode = pptInfoType.oName
                    Case "Abkürzung"
                        catCode = pptInfoType.sName
                    Case "voller Name"
                        catCode = pptInfoType.bCrumb
                    Case "Lieferumfänge"
                        catCode = pptInfoType.lUmfang
                    Case "manuelle Termin-Änderungen"
                        catCode = pptInfoType.mvElement
                    Case "Ressourcen"
                        catCode = pptInfoType.resources
                    Case "Kosten"
                        catCode = pptInfoType.costs
                    Case "Überfällig"
                        catCode = pptInfoType.overDue
                    Case Else
                        catCode = pptInfoType.cName
                End Select
            End If


            'If rdbName.Checked Then
            '    rdbCode = pptInfoType.cName
            'ElseIf rdbOriginalName.Checked Then
            '    rdbCode = pptInfoType.oName
            'ElseIf rdbAbbrev.Checked Then
            '    rdbCode = pptInfoType.sName
            'ElseIf rdbBreadcrumb.Checked Then
            '    rdbCode = pptInfoType.bCrumb
            'ElseIf rdbLU.Checked Then
            '    rdbCode = pptInfoType.lUmfang
            'ElseIf rdbMV.Checked Then
            '    rdbCode = pptInfoType.mvElement
            'ElseIf rdbResources.Checked Then
            '    rdbCode = pptInfoType.resources
            'ElseIf rdbCosts.Checked Then
            '    rdbCode = pptInfoType.costs
            'Else
            '    rdbCode = pptInfoType.cName
            'End If

            Dim nameCollection As Collection

            If selectedLanguage <> defaultSprache And catCode = pptInfoType.cName Then
                If suchString = "" Then
                    nameCollection = smartSlideLists.getNCollection(colorCode, suchString, catCode)
                    ' jetzt müssen die Namen in NameCollection erstmal ersetzt werden 
                    Dim tmpCollection As New Collection
                    For Each elemName As String In nameCollection
                        Dim newName As String = languages.translate(elemName, selectedLanguage)
                        ' es ist sichergestellt, dass es keine Doubletten gibt, also jedes Wort kann eindeutig übersetzt werden 
                        If Not tmpCollection.Contains(newName) Then
                            tmpCollection.Add(newName, newName)
                        End If
                    Next
                    nameCollection.Clear()
                    nameCollection = tmpCollection
                Else
                    ' jetzt müssen die anders-sprachigen Namen erstmal mit dem suchstring gefiltert werden 
                    Dim tmpCollection As New Collection
                    For Each anderName As String In Me.listboxNames.Items
                        If anderName.Contains(suchString) Then
                            If Not tmpCollection.Contains(anderName) Then
                                tmpCollection.Add(anderName, anderName)
                            End If
                        End If
                    Next

                    ' dann müssen die anders-sprachigen Namen in die Original Namen übersetzt und per Farb-Code gefiltert werden 
                    Dim oNameCollection As New Collection
                    For Each anderName As String In tmpCollection
                        Dim newName As String = languages.backtranslate(anderName, selectedLanguage)
                        If Not oNameCollection.Contains(newName) Then
                            oNameCollection.Add(newName, newName)
                        End If
                    Next

                    ' jetzt nach Farbcode ausdünnen ...
                    'If colorCode = 0 Or colorCode = 15 Then
                    '    oNameCollection = smartSlideLists.getTNCollection(colorCode, oNameCollection)
                    'End If
                    ' das vorherige war doch falsch ... weil ja dann gar nichts aussortiert wurde ... 
                    oNameCollection = smartSlideLists.getTNCollection(colorCode, oNameCollection)


                    ' was jetzt übrig bleibt, muss wieder in die Ander-Sprache zurückkonvertiert werden 
                    ' dann müssen die anders-sprachigen Namen in die Original Namen übersetzt und per Farb-Code gefiltert werden 
                    nameCollection = New Collection
                    For Each oName As String In oNameCollection
                        Dim newName As String = languages.translate(oName, selectedLanguage)
                        If Not nameCollection.Contains(newName) Then
                            nameCollection.Add(newName, newName)
                        End If
                    Next


                End If

            Else
                nameCollection = smartSlideLists.getNCollection(colorCode, suchString, catCode)
            End If

            ' die bisherige Liste zurücksetzen
            Me.listboxNames.Items.Clear()

            For Each elem As Object In nameCollection
                listboxNames.Items.Add(CStr(elem))
            Next

            'listboxNames.Focus()

        End If

    End Sub


    Private Sub ucSearch_Load(sender As Object, e As EventArgs) Handles Me.Load

        cathegoryList.MaxDropDownItems = 9
        If englishLanguage Then
            With Me
                .Label1.Text = "Search Results:"
                .Label2.Text = "Elements:"
            End With
            cathegoryList.Items.Add("Name")
            cathegoryList.Items.Add("Responsibilities")
            cathegoryList.Items.Add("Deliverables")
            'cathegoryList.Items.Add("Original Name")
            cathegoryList.Items.Add("Overdue")
            cathegoryList.Items.Add("Resources")
            cathegoryList.Items.Add("Abbreviation")
            cathegoryList.Items.Add("Cost")
            cathegoryList.Items.Add("manually Changed Dates")

            'If slideHasSmartElements Then
            '    cathegoryList.SelectedItem = "Name"
            'End If
        Else
            With Me
                .Label1.Text = "Suchergebnisse:"
                .Label2.Text = "Elemente:"
            End With
            cathegoryList.Items.Add("Name")
            cathegoryList.Items.Add("Verantwortlich")
            cathegoryList.Items.Add("Lieferumfänge")
            'cathegoryList.Items.Add("Original Name")
            cathegoryList.Items.Add("Überfällig")
            cathegoryList.Items.Add("Ressourcen")
            cathegoryList.Items.Add("Abkürzung")
            cathegoryList.Items.Add("Kosten")
            cathegoryList.Items.Add("manuelle Termin-Änderungen")

        End If

        'Call fülltListbox()

    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub filterText_MouseHover(sender As Object, e As EventArgs) Handles filterText.MouseHover
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "Text-filter for listbox"
        Else
            tsMSG = "Text-Filter für Listbox"
        End If
        'ToolTip1.Show(tsMSG, filterText, 2000)
    End Sub

    Private Sub filterText_TextChanged(sender As Object, e As EventArgs) Handles filterText.TextChanged

        ' ''selListboxNames.Items.Clear()
        suchString = filterText.Text
        Call fülltListbox(hShowTrafficLights)

    End Sub



    Private Sub listboxNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listboxNames.SelectedIndexChanged

        ' es werden alle selektierten Namen als Shapes selektiert ....
        ' pro Name auch mehrere Shapes selektiert werden müssen 
        ' wenn Ampeln anzeigen an ist, dann werden auch die entsprechenden Ampel-Farben angezeigt ... 

        ' die zur Auswahl gehörenden Namen werden gelöscht 
        selListboxNames.Items.Clear()

        Dim nameArrayI() As String
        Dim nameArrayO() As String
        Dim anzSelected As Integer = listboxNames.SelectedItems.Count

        Dim catCode As Integer

        If englishLanguage Then
            Select Case cathegoryList.SelectedItem
                Case "Name"
                    catCode = pptInfoType.cName
                Case "Responsibilities"
                    catCode = pptInfoType.responsible
                Case "Original Name"
                    catCode = pptInfoType.oName
                Case "Abbreviation"
                    catCode = pptInfoType.sName
                Case "voller Name"
                    catCode = pptInfoType.bCrumb
                Case "Deliverables"
                    catCode = pptInfoType.lUmfang
                Case "manually Changed Dates"
                    catCode = pptInfoType.mvElement
                Case "Resources"
                    catCode = pptInfoType.resources
                Case "Cost"
                    catCode = pptInfoType.costs
                Case "Overdue"
                    catCode = pptInfoType.overDue
                Case Else
                    catCode = pptInfoType.cName
            End Select
        Else
            Select Case cathegoryList.SelectedItem
                Case "Name"
                    catCode = pptInfoType.cName
                Case "Verantwortlich"
                    catCode = pptInfoType.responsible
                Case "Original Name"
                    catCode = pptInfoType.oName
                Case "Abkürzung"
                    catCode = pptInfoType.sName
                Case "voller Name"
                    catCode = pptInfoType.bCrumb
                Case "Lieferumfänge"
                    catCode = pptInfoType.lUmfang
                Case "manuelle Termin-Änderungen"
                    catCode = pptInfoType.mvElement
                Case "Ressourcen"
                    catCode = pptInfoType.resources
                Case "Kosten"
                    catCode = pptInfoType.costs
                Case "Überfällig"
                    catCode = pptInfoType.overDue
                Case Else
                    catCode = pptInfoType.cName
            End Select
        End If

        

        ' ''If rdbName.Checked Then
        ' ''    rdbCode = pptInfoType.cName
        ' ''ElseIf rdbOriginalName.Checked Then
        ' ''    rdbCode = pptInfoType.oName
        ' ''ElseIf rdbAbbrev.Checked Then
        ' ''    rdbCode = pptInfoType.sName
        ' ''ElseIf rdbBreadcrumb.Checked Then
        ' ''    rdbCode = pptInfoType.bCrumb
        ' ''ElseIf rdbLU.Checked Then
        ' ''    rdbCode = pptInfoType.lUmfang
        ' ''ElseIf rdbMV.Checked Then
        ' ''    rdbCode = pptInfoType.mvElement
        ' ''ElseIf rdbResources.Checked Then
        ' ''    rdbCode = pptInfoType.resources
        ' ''ElseIf rdbCosts.Checked Then
        ' ''    rdbCode = pptInfoType.costs
        ' ''Else
        ' ''    rdbCode = pptInfoType.cName
        ' ''End If

        ReDim nameArrayI(anzSelected - 1)

        For i As Integer = 0 To anzSelected - 1
            Dim tmpText As String = CStr(listboxNames.SelectedItems.Item(i))

            ' jetzt muss gechecked werden, ob noch übersetzt werden muss
            If catCode = pptInfoType.cName And selectedLanguage <> defaultSprache Then
                tmpText = languages.backtranslate(tmpText, selectedLanguage)
            End If

            nameArrayI(i) = tmpText
        Next

        Dim colorCode As Integer = calcColorCode(hShowTrafficLights)

        ' in tmpCollection werden alle Elementnamen geschoben, auf die die Selection in listboxNames zutrifft

        Dim tmpCollection As Collection = smartSlideLists.getShapesNames(nameArrayI, catCode, colorCode)

        anzSelected = tmpCollection.Count


        If anzSelected >= 1 Then

            ' wenn das erste Element selektiert wird und die Anzahl Marker > 0 ist, dann müssen hier die MArker gelöscht werden 
            If listboxNames.SelectedItems.Count = 1 And markerShpNames.Count > 0 Then
                Call deleteMarkerShapes()
            End If

            ReDim nameArrayO(anzSelected - 1)

            For i As Integer = 0 To anzSelected - 1
                nameArrayO(i) = CStr(tmpCollection.Item(i + 1))
            Next

            Try
                selectedPlanShapes = currentSlide.Shapes.Range(nameArrayO)

                For Each tmpshape As PowerPoint.Shape In selectedPlanShapes
                    If isProjectCardInvisible(tmpshape) Then
                        tmpshape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    End If
                Next

                ' wird ganz am Ende gemacht ...
                'selectedPlanShapes.Select()

                ' die WindowsSelection Change Routine gleich wieder verlassen ... damit die MArkerShapes nicht gleich wieder gelöscht werden 

                If showMarker Then
                    If selectedPlanShapes.Count > 1 Then

                        Call createMarkerShapes(pptShapes:=selectedPlanShapes)

                    ElseIf selectedPlanShapes.Count = 1 Then

                        Call createMarkerShapes(pptShape:=selectedPlanShapes.Item(1))

                    End If
                End If


                ' löschen der Liste an Elemente, die zur Selection gehören
                selListboxNames.Items.Clear()
                shpNameSav.Clear()

                Dim selListboxEle As String = ""

                'neue Elemente aus Selection in Liste bringen
                For Each selEleShpName In tmpCollection

                    Dim curShape As PowerPoint.Shape = currentSlide.Shapes(selEleShpName)

                    'Dim bln As String = curShape.Tags.Item("BLN")
                    Dim bln As String = bestimmeElemText(curShape, False, False, showBestName)
                    Dim pname As String = ""
                    If isProjectCard(curShape) Then
                        pname = getPVnameFromTags(selEleShpName)
                    Else
                        pname = getPVnameFromShpName(selEleShpName)
                    End If

                    selListboxEle = pname & "  -  " & bln
                    ' merken der Zuordnung zwischen angezeigtem Namen und ShapeNamen

                    Dim lfdNr As Integer = 1
                    Dim tmpKey As String = selListboxEle
                    Do While shpNameSav.ContainsKey(tmpKey)
                        lfdNr = lfdNr + 1
                        tmpKey = selListboxEle & " (" & lfdNr.ToString & ")"
                    Loop
                    shpNameSav.Add(tmpKey, selEleShpName)

                    selListboxNames.Items.Add(tmpKey)

                Next

                selectedPlanShapes.Select()

            Catch ex As Exception

            End Try



        Else
            ' nichts tun ...

        End If


    End Sub

    Private Sub cathegoryList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cathegoryList.SelectedIndexChanged

        Dim newCathegory As String = cathegoryList.SelectedItem
        Dim oldFilterText As String = filterText.Text

        filterText.Text = ""
        ' tk, 11.1.18 das folgende muss nur dann aufgerufen werden, wenn es keine Änderung im Filtertext-Feld gab. Dann muss das fülltListbox  in dem Event Aufruf von cathegoryList_SelectedIndexChanged
        If oldFilterText = "" Then
            Call fülltListbox(hShowTrafficLights)
        End If

        listboxNames.Focus()

    End Sub

    Private Sub selListboxNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles selListboxNames.SelectedIndexChanged

        Dim nameArrayO() As String
        Dim anzSelected As Integer = selListboxNames.SelectedItems.Count
        Dim ShpName As String = ""

        If anzSelected = 1 Then


            ' wenn das erste Element selektiert wird und die Anzahl Marker > 0 ist, dann müssen hier die MArker gelöscht werden 
            If selListboxNames.SelectedItems.Count = 1 And markerShpNames.Count > 0 Then
                Call deleteMarkerShapes()
            End If

            ReDim nameArrayO(anzSelected - 1)

            ShpName = CStr(shpNameSav.Item(selListboxNames.SelectedItem))
            nameArrayO(0) = ShpName
           
            Try
                selectedPlanShapes = currentSlide.Shapes.Range(nameArrayO)
                selectedPlanShapes.Select()

                ' die WindowsSelection Change Routine gleich wieder verlassen ... damit die MArkerShapes nicht gleich wieder gelöscht werden 

                If showMarker Then
                    If selectedPlanShapes.Count > 1 Then

                        Call createMarkerShapes(pptShapes:=selectedPlanShapes)

                    ElseIf selectedPlanShapes.Count = 1 Then

                        Call createMarkerShapes(pptShape:=selectedPlanShapes.Item(1))

                    End If
                End If

            Catch ex As Exception

            End Try

        Else
            ' nichts tun ...

        End If


        Dim curShape As PowerPoint.Shape = currentSlide.Shapes(ShpName)

        Call aktualisiereInfoPane(curShape, False)

    End Sub

    Private Sub listboxNames_SelectedValueChanged(sender As Object, e As EventArgs) Handles listboxNames.SelectedValueChanged

    End Sub

    Private Sub selListboxNames_SelectedValueChanged(sender As Object, e As EventArgs) Handles selListboxNames.SelectedValueChanged

    End Sub
End Class
