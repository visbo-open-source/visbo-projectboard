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

    ' wird in den entsprechenden Checkbox Routinen gesetzt 
    Private colorCode As Integer = 0

    ' wird im entsprechenden Suchfeld gesetzt 
    Private suchString As String = ""


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

    ''' <summary>
    ''' füllt die ListboxNames mit den Elementen, deren Ampel keine Bewertung hat
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub shwOhneLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwOhneLight.CheckedChanged

        Dim ampelColor As Integer = PTfarbe.none
        showTrafficLights(ampelColor) = shwOhneLight.Checked

        If shwOhneLight.Checked Then

        End If

        Call fülltListbox()

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
        showTrafficLights(ampelColor) = shwGreenLight.Checked

        If shwGreenLight.Checked Then

        End If

        Call fülltListbox()

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
        showTrafficLights(ampelColor) = shwYellowLight.Checked

        If shwYellowLight.Checked Then

        End If

        Call fülltListbox()

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
        showTrafficLights(ampelColor) = shwRedLight.Checked

        If shwRedLight.Checked Then

        End If

        Call fülltListbox()

        Call faerbeShapes(ampelColor, shwRedLight.Checked)
    End Sub

    ''' <summary>
    ''' erstellt die Listbox aufgrund der Settings bei Ampeln, Radio-Button und Suchstr neu 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub fülltListbox()

        If Not dontFire Then

            colorCode = calcColorCode()

            Dim catCode As Integer


            ''
            '' hier muss die Textbox categoryList ausgelesen werden. Hier wird gefiltert
            ''

            Select Case cathegoryList.SelectedItem
                Case "Name"
                    catCode = pptInfoType.cName
                Case "Original Name"
                    catCode = pptInfoType.oName
                Case "Abkürzung"
                    catCode = pptInfoType.sName
                Case "voller Name"
                    catCode = pptInfoType.bCrumb
                Case "Lieferumfänge"
                    catCode = pptInfoType.lUmfang
                Case "Termin-Änderungen"
                    catCode = pptInfoType.mvElement
                Case "Ressourcen"
                    catCode = pptInfoType.resources
                Case "Kosten"
                    catCode = pptInfoType.costs
                Case Else
                    catCode = pptInfoType.cName
            End Select

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
        End If

    End Sub


    Private Sub ucSearch_Load(sender As Object, e As EventArgs) Handles Me.Load

        cathegoryList.MaxDropDownItems = 8
        cathegoryList.Items.Add("Name")
        cathegoryList.Items.Add("Lieferumfänge")
        cathegoryList.Items.Add("Original Name")
        cathegoryList.Items.Add("Ressourcen")
        cathegoryList.Items.Add("Abkürzung")
        cathegoryList.Items.Add("Kosten")
        cathegoryList.Items.Add("Termin-Änderungen")
        cathegoryList.SelectedItem = "Name"

        Call fülltListbox()

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

        suchString = filterText.Text
        Call fülltListbox()

    End Sub

    Private Sub listboxNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listboxNames.SelectedIndexChanged

        ' es werden alle selektierten Namen als Shapes selektiert ....
        ' es können pro Name auch mehrere Shapes selektiert werden müssen 
        ' wenn Ampeln anzeigen an ist, dann werden auch die entsprechenden Ampel-Farben angezeigt ... 


        Dim nameArrayI() As String
        Dim nameArrayO() As String
        Dim anzSelected As Integer = listboxNames.SelectedItems.Count

        Dim catCode As Integer


        Select Case cathegoryList.SelectedItem
            Case "Name"
                catCode = pptInfoType.cName
            Case "Original Name"
                catCode = pptInfoType.oName
            Case "Abkürzung"
                catCode = pptInfoType.sName
            Case "voller Name"
                catCode = pptInfoType.bCrumb
            Case "Lieferumfänge"
                catCode = pptInfoType.lUmfang
            Case "Termin-Änderungen"
                catCode = pptInfoType.mvElement
            Case "Ressourcen"
                catCode = pptInfoType.resources
            Case "Kosten"
                catCode = pptInfoType.costs
            Case Else
                catCode = pptInfoType.cName
        End Select

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

        Dim colorCode As Integer = calcColorCode()

        Dim tmpCollection As Collection = smartSlideLists.getShapesNames(nameArrayI, catCode, colorCode)

        anzSelected = tmpCollection.Count


        If anzSelected >= 1 Then

            ' wenn das erste Element selektiert wird udn die Anzahl Marker > 0 ist, dann müssen hier die MArker gelöscht werden 
            If listboxNames.SelectedItems.Count = 1 And markerShpNames.Count > 0 Then
                Call deleteMarkerShapes()
            End If

            ReDim nameArrayO(anzSelected - 1)

            For i As Integer = 0 To anzSelected - 1
                nameArrayO(i) = CStr(tmpCollection.Item(i + 1))
            Next

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


    End Sub

    Private Sub cathegoryList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cathegoryList.SelectedIndexChanged

        Dim newCathegory As String = cathegoryList.SelectedItem

        Call fülltListbox()

    End Sub
End Class
