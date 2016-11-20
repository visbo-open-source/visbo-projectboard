''' <summary>
''' das Form Info wird in variabler Größe angezeigt: mit / ohne Ampel-Block, mit /ohne Search-Block
''' es gibt zwei Methoden ampelblockVisibible und searchblockVisible, die die Elemente dann entsprechend positionieren und sichtbar machen 
''' </summary>
''' <remarks></remarks>
Public Class frmInfo

    Friend abkuerzung As String
    Friend showSearchListBox As Boolean = False

    Private Const deltaAmpel As Integer = 50
    Private Const deltaSearchBox As Integer = 200
    Private Const smallHeight As Integer = 220

    Private dontFire As Boolean = False
    ' innerhalb der Klasse überall im Zugriff; Colorcode ist die Zahl , die sich ergibt , 
    ' wenn man die Werte 0, 1, 2, 3 als Potenzen von 2 und in Summe ausrechnet

    ' wird in den entsprechenden Checkbox Routinen gesetzt 
    Private colorCode As Integer = 0

    ' steuert, wo der Text relatic zum Meilenstein , zur Phase platziert werden soll 
    ' MD: MilestoneDate, MT MilestoneText , PD PhaseDate, PT PhaseText
    Friend positionIndexMD As Integer = 5
    Friend positionIndexMT As Integer = 1
    Friend positionIndexPD As Integer = 8
    Friend positionIndexPT As Integer = 6

    ' wird im entsprechenden Suchfeld gesetzt 
    Private suchString As String = ""

    Private Sub frmInfo_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        infoFrm = Nothing
        formIsShown = False
    End Sub

    Private Sub frmInfo_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        formIsShown = False
    End Sub

    ''' <summary>
    ''' zeigt den Ampel-/LU-/Moved Erläuterungstext inkl de rbutton und verschiebt die anderen Elemente entsprechend 
    ''' Ändert die Höhen von TabControl1 und des gesamten Formulars  
    ''' </summary>
    ''' <param name="istSichtbar"></param>
    ''' <remarks></remarks>
    Private Sub aLuTvBlockVisible(ByVal istSichtbar As Boolean)

        ' Größen und Positionen anpassen 
        If Not istSichtbar Then
            With Me
                .Height = Me.Height - deltaAmpel
                .TabControl1.Height = Me.TabControl1.Height - deltaAmpel
                .filterText.Top = .filterText.Top - deltaAmpel
                .searchIcon.Top = .searchIcon.Top - deltaAmpel
                .btnSendToHome.Top = .btnSendToHome.Top - deltaAmpel
                .btnSentToChange.Top = .btnSentToChange.Top - deltaAmpel
                '.PictureMarker.Top = .PictureMarker.Top - deltaAmpel
                '.CheckBxMarker.Top = .CheckBxMarker.Top - deltaAmpel
                .listboxNames.Top = .listboxNames.Top - deltaAmpel
                .rdbName.Top = .rdbName.Top - deltaAmpel
                .rdbLU.Top = .rdbLU.Top - deltaAmpel
                .rdbMV.Top = .rdbMV.Top - deltaAmpel
                .rdbOriginalName.Top = rdbOriginalName.Top - deltaAmpel
                .rdbAbbrev.Top = rdbAbbrev.Top - deltaAmpel
                .rdbBreadcrumb.Top = .rdbBreadcrumb.Top - deltaAmpel
            End With
        Else
            With Me
                .Height = Me.Height + deltaAmpel
                .TabControl1.Height = Me.TabControl1.Height + deltaAmpel
                .filterText.Top = .filterText.Top + deltaAmpel
                .searchIcon.Top = .searchIcon.Top + deltaAmpel
                .btnSendToHome.Top = .btnSendToHome.Top + deltaAmpel
                .btnSentToChange.Top = .btnSentToChange.Top + deltaAmpel
                '.PictureMarker.Top = .PictureMarker.Top + deltaAmpel
                '.CheckBxMarker.Top = .CheckBxMarker.Top + deltaAmpel
                .listboxNames.Top = .listboxNames.Top + deltaAmpel
                .rdbName.Top = .rdbName.Top + deltaAmpel
                .rdbLU.Top = .rdbLU.Top + deltaAmpel
                .rdbMV.Top = .rdbMV.Top + deltaAmpel
                .rdbOriginalName.Top = rdbOriginalName.Top + deltaAmpel
                .rdbAbbrev.Top = rdbAbbrev.Top + deltaAmpel
                .rdbBreadcrumb.Top = .rdbBreadcrumb.Top + deltaAmpel
            End With

        End If

        Me.aLuTvText.Visible = istSichtbar
        Me.deleteAmpel.Visible = istSichtbar
        Me.writeAmpel.Visible = istSichtbar

    End Sub

    ''' <summary>
    ''' zeigt die Searchbox an bzw. macht sie unsichtbar
    ''' verändert die Größen des Formulars entsprechend 
    ''' </summary>
    ''' <param name="istSichtbar"></param>
    ''' <remarks></remarks>
    Private Sub searchBlockVisible(ByVal istSichtbar As Boolean)

        If Not istSichtbar Then
            ' es soll nicht sichtbar sein 
            Me.Height = Me.Height - deltaSearchBox
            filterText.Visible = False
            listboxNames.Visible = False
            rdbName.Visible = False
            rdbLU.Visible = False
            rdbMV.Visible = False
            rdbOriginalName.Visible = False
            rdbAbbrev.Visible = False
            rdbBreadcrumb.Visible = False
        Else
            Me.Height = Me.Height + deltaSearchBox
            filterText.Visible = True
            listboxNames.Visible = True
            rdbName.Visible = True
            rdbLU.Visible = True
            rdbMV.Visible = True
            If extSearch Then
                rdbOriginalName.Visible = True
                rdbAbbrev.Visible = True
                rdbBreadcrumb.Visible = True
            End If
        End If

    End Sub


    Private Sub frmInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' sind irgendwelche Ampel-Farben gesetzt 
        Dim ix As Integer = 1

        formIsShown = True

        PictureMarker.Visible = True
        CheckBxMarker.Visible = True

        Do While ix <= 3 And Not ampelnExistieren
            Dim tmpCollection As Collection = smartSlideLists.getShapeNamesWithColor(ix)
            If tmpCollection.Count > 0 Then
                ampelnExistieren = True
            Else
                ix = ix + 1
            End If

        Loop

        If ampelnExistieren Then
            With Me.shwGreenLight
                .Checked = False
                .Visible = True
            End With

            With Me.shwYellowLight
                .Checked = False
                .Visible = True
            End With

            With Me.shwRedLight
                .Checked = False
                .Visible = True
            End With

            With Me.shwOhneLight
                .Checked = False
                .Visible = True
            End With

            With Me.lblAmpeln
                .Visible = True
            End With
        Else

            With Me.shwGreenLight
                .Checked = False
                .Visible = False
            End With

            With Me.shwYellowLight
                .Checked = False
                .Visible = False
            End With

            With Me.shwRedLight
                .Checked = False
                .Visible = False
            End With

            With Me.shwOhneLight
                .Checked = False
                .Visible = False
            End With

            With Me.lblAmpeln
                .Visible = False
            End With
        End If

        CheckBxMarker.Checked = showMarker
        ' Zu Beginn ist Ampel-Text und Ampel-Erläuterung nicht sichtbar 
        Call aLuTvBlockVisible(False)

        ' Zu Beginn ist die Searchbox nicht visible 
        Call searchBlockVisible(False)

        dontFire = True

        showOrginalName.Visible = False
        showOrigName = False

        If showBreadCrumbField = True Then
            fullBreadCrumb.Visible = True
        Else
            fullBreadCrumb.Visible = False
        End If


        showAbbrev.Checked = showShortName

        ' jetzt muss geprüft werden, ob GoToHome und GoToChangedPos enabled sind ... 
        btnSentToChange.Enabled = changedButtonRelevance
        btnSendToHome.Enabled = homeButtonRelevance

        dontFire = False

        ' ab jetzt sollen wieder die entsprechenden Event Routinen durchgeführt werden 
        With Me.rdbName
            .Checked = True
        End With

    End Sub

    Private Sub rdbName_CheckedChanged(sender As Object, e As EventArgs) Handles rdbName.CheckedChanged
        ' dontFire true verhindert, dass die Aktion durchgeführt wird, das ist dann erforderlich wenn man explizit verhindern will, 
        ' dass ständig die Events getriggert werden 


        If rdbName.Checked = True Then
            Me.aLuTvText.Text = setALuTvText


            Call erstelleListbox()

        End If
    End Sub

    ''' <summary>
    ''' erstellt die Listbox aufgrund der Settings bei Ampeln, Radio-Button und Suchstr neu 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub erstelleListbox()

        If Not dontFire Then

            colorCode = calcColorCode()

            Dim rdbCode As Integer

            If rdbName.Checked Then
                rdbCode = pptInfoType.cName
            ElseIf rdbOriginalName.Checked Then
                rdbCode = pptInfoType.oName
            ElseIf rdbAbbrev.Checked Then
                rdbCode = pptInfoType.sName
            ElseIf rdbBreadcrumb.Checked Then
                rdbCode = pptInfoType.bCrumb
            ElseIf rdbLU.Checked Then
                rdbCode = pptInfoType.lUmfang
            ElseIf rdbMV.Checked Then
                rdbCode = pptInfoType.mvElement
            Else
                rdbCode = pptInfoType.cName
            End If

            Dim nameCollection As Collection

            If selectedLanguage <> defaultSprache And rdbCode = pptInfoType.cName Then
                If suchString = "" Then
                    nameCollection = smartSlideLists.getNCollection(colorCode, suchString, rdbCode)
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
                nameCollection = smartSlideLists.getNCollection(colorCode, suchString, rdbCode)
            End If

            ' die bisherige Liste zurücksetzen
            Me.listboxNames.Items.Clear()

            For Each elem As Object In nameCollection
                listboxNames.Items.Add(CStr(elem))
            Next
        End If

    End Sub

    ''' <summary>
    ''' berechnet eine Integer Zahl, die Auskunft gibt, wie die vier Checkboxen gesetzt sind 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function calcColorCode() As Integer

        Dim tmpNumber As Integer = 0

        If Not ampelnExistieren Then
            tmpNumber = 0
        Else
            If Me.shwOhneLight.Checked Then
                tmpNumber = tmpNumber + 1 ' 2 hoch 0 
            End If

            If Me.shwGreenLight.Checked Then
                tmpNumber = tmpNumber + 2 ' 2 hoch 1 
            End If

            If Me.shwYellowLight.Checked Then
                tmpNumber = tmpNumber + 4 ' 2 hoch 2 
            End If

            If Me.shwRedLight.Checked Then
                tmpNumber = tmpNumber + 8 ' 2 hoch 3 
            End If
        End If

        calcColorCode = tmpNumber

    End Function

    ''' <summary>
    ''' zeigt bei den Shapes, die die angegebene Ampelfarbe haben, diese Farbe als Hintergrund Schatten an bzw. löscht den Hintergrund Schatten wieder
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <param name="show"></param>
    ''' <remarks></remarks>
    Private Sub faerbeShapes(ByVal ampelColor As Integer, ByVal show As Boolean)

        Dim tmpCollection As Collection = smartSlideLists.getShapeNamesWithColor(ampelColor)
        Dim anzSelected As Integer = tmpCollection.Count
        Dim nameArray() As String

        If ampelColor >= 0 And ampelColor <= 3 Then
            'alles ok 
        Else
            ' sicherstellen, es kommt zu keinem Absturz .... 
            ampelColor = 0
        End If

        Dim farben(4) As Long
        farben(0) = PowerPoint.XlRgbColor.rgbGrey
        farben(1) = PowerPoint.XlRgbColor.rgbGreen
        farben(2) = PowerPoint.XlRgbColor.rgbYellow
        farben(3) = PowerPoint.XlRgbColor.rgbRed
        farben(4) = PowerPoint.XlRgbColor.rgbWhite

        Dim shapesToBeColored As PowerPoint.ShapeRange

        If anzSelected >= 1 Then
            ReDim nameArray(anzSelected - 1)

            For i As Integer = 0 To anzSelected - 1
                nameArray(i) = CStr(tmpCollection.Item(i + 1))
            Next

            Try
                shapesToBeColored = currentSlide.Shapes.Range(nameArray)

                If show Then
                    ' mit Schatten einfärben 
                    With shapesToBeColored.Shadow
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        .Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
                        .Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                        .Blur = 0
                        .Size = 160
                        .OffsetX = 0
                        .OffsetY = 0
                        .Transparency = 0
                        .ForeColor.RGB = farben(ampelColor)
                    End With
                Else
                    ' Schatten wieder wegnehmen 
                    With shapesToBeColored.Shadow
                        .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    End With
                End If


            Catch ex As Exception

            End Try

        Else
            ' nichts tun ...

        End If


    End Sub

    Private Sub rdbOriginalName_CheckedChanged(sender As Object, e As EventArgs) Handles rdbOriginalName.CheckedChanged
        If rdbOriginalName.Checked = True Then
            Me.aLuTvText.Text = setALuTvText

            Call erstelleListbox()

        End If
    End Sub

    Private Sub rdbAbbrev_CheckedChanged(sender As Object, e As EventArgs) Handles rdbAbbrev.CheckedChanged
        If rdbAbbrev.Checked = True Then
            Me.aLuTvText.Text = setALuTvText
            Call erstelleListbox()

        End If
    End Sub

    Private Sub rdbBreadcrumb_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBreadcrumb.CheckedChanged

        If rdbBreadcrumb.Checked = True Then

            Me.aLuTvText.Text = setALuTvText
            
            Call erstelleListbox()

        End If

    End Sub

    ''' <summary>
    ''' bestimmt den String in Abhängigkeit von rdbCode und dem selektierten Shape 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function setALuTvText() As String

        Dim tmpResult As String
        If Not IsNothing(selectedPlanShapes) Then
            If selectedPlanShapes.Count = 1 Then
                Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                Dim rdbcode = calcRDB()
                tmpResult = bestimmeElemALuTvText(tmpShape, rdbcode)
            Else
                tmpResult = ""
            End If
        Else
            tmpResult = ""
        End If

        setALuTvText = tmpResult
    End Function
    Private Sub listboxNames_DoubleClick(sender As Object, e As EventArgs) Handles listboxNames.DoubleClick
        If rdbMV.Checked = True Then
            ' jetzt kann der Erläuterungstext eingegeben werden ... 
            Call MsgBox("Erläuterung eingeben ...")
        End If
    End Sub

    Private Sub listboxNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listboxNames.SelectedIndexChanged

        ' es werden alle selektierten Namen als Shapes selektiert ....
        ' es können pro Name auch mehrere Shapes selektiert werden müssen 
        ' wenn Ampeln anzeigen an ist, dann werden auch die entsprechenden Ampel-Farben angezeigt ... 


        Dim nameArrayI() As String
        Dim nameArrayO() As String
        Dim anzSelected As Integer = listboxNames.SelectedItems.Count

        Dim rdbCode As Integer

        If rdbName.Checked Then
            rdbCode = pptInfoType.cName
        ElseIf rdbOriginalName.Checked Then
            rdbCode = pptInfoType.oName
        ElseIf rdbAbbrev.Checked Then
            rdbCode = pptInfoType.sName
        ElseIf rdbBreadcrumb.Checked Then
            rdbCode = pptInfoType.bCrumb
        ElseIf rdbLU.Checked Then
            rdbCode = pptInfoType.lUmfang
        ElseIf rdbMV.Checked Then
            rdbCode = pptInfoType.mvElement
        Else
            rdbCode = pptInfoType.cName
        End If

        ReDim nameArrayI(anzSelected - 1)

        For i As Integer = 0 To anzSelected - 1
            Dim tmpText As String = CStr(listboxNames.SelectedItems.Item(i))

            ' jetzt muss gechecked werden, ob noch übersetzt werden muss
            If rdbCode = pptInfoType.cName And selectedLanguage <> defaultSprache Then
                tmpText = languages.backtranslate(tmpText, selectedLanguage)
            End If

            nameArrayI(i) = tmpText
        Next

        Dim colorCode As Integer = calcColorCode()

        Dim tmpCollection As Collection = smartSlideLists.getShapesNames(nameArrayI, rdbCode, colorCode)

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

    Private Sub filterText_TextChanged(sender As Object, e As EventArgs) Handles filterText.TextChanged
        suchString = filterText.Text
        Call erstelleListbox()
    End Sub

    Private Sub shwOhneLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwOhneLight.CheckedChanged

        If shwOhneLight.Checked Then
            If Not Me.aLuTvText.Visible Then
                Call aLuTvBlockVisible(True)
            End If
        Else
            If Me.aLuTvText.Visible And _
                    (Not shwGreenLight.Checked And Not shwYellowLight.Checked And Not shwRedLight.Checked) And _
                    (Not Me.rdbLU.Checked And Not Me.rdbMV.Checked) Then
                Call aLuTvBlockVisible(False)
            End If
        End If

        Call erstelleListbox()
        Dim ampelColor As Integer = 0
        Call faerbeShapes(ampelColor, shwOhneLight.Checked)
    End Sub

    Private Sub shwGreenLight_CheckedChanged_1(sender As Object, e As EventArgs) Handles shwGreenLight.CheckedChanged

        If shwGreenLight.Checked Then
            If Not Me.aLuTvText.Visible Then
                Call aLuTvBlockVisible(True)
            End If
        Else
            If Me.aLuTvText.Visible And _
                    (Not shwOhneLight.Checked And Not shwYellowLight.Checked And Not shwRedLight.Checked) And _
                    (Not Me.rdbLU.Checked And Not Me.rdbMV.Checked) Then
                Call aLuTvBlockVisible(False)
            End If
        End If

        Call erstelleListbox()
        Dim ampelColor As Integer = 1
        Call faerbeShapes(ampelColor, shwGreenLight.Checked)
    End Sub

    Private Sub shwYellowLight_CheckedChanged_1(sender As Object, e As EventArgs) Handles shwYellowLight.CheckedChanged

        If shwYellowLight.Checked Then
            If Not Me.aLuTvText.Visible Then
                Call aLuTvBlockVisible(True)
            End If
        Else
            If Me.aLuTvText.Visible And _
                    (Not shwGreenLight.Checked And Not shwOhneLight.Checked And Not shwRedLight.Checked) And _
                    (Not Me.rdbLU.Checked And Not Me.rdbMV.Checked) Then
                Call aLuTvBlockVisible(False)
            End If
        End If

        Call erstelleListbox()
        Dim ampelColor As Integer = 2
        Call faerbeShapes(ampelColor, shwYellowLight.Checked)
    End Sub

    Private Sub shwRedLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwRedLight.CheckedChanged

        If shwRedLight.Checked Then
            If Not Me.aLuTvText.Visible Then
                Call aLuTvBlockVisible(True)
            End If
        Else
            If Me.aLuTvText.Visible And _
                    (Not shwGreenLight.Checked And Not shwOhneLight.Checked And Not shwYellowLight.Checked) And _
                    (Not Me.rdbLU.Checked And Not Me.rdbMV.Checked) Then

                Call aLuTvBlockVisible(False)

            End If
        End If

        Call erstelleListbox()
        Dim ampelColor As Integer = 3
        Call faerbeShapes(ampelColor, shwRedLight.Checked)
    End Sub

    Private Sub showAbbrev_CheckedChanged(sender As Object, e As EventArgs) Handles showAbbrev.CheckedChanged

        showShortName = showAbbrev.Checked

        If dontFire Then
            Exit Sub
        End If

        Try
            If showAbbrev.Checked Then

                dontFire = True
                Me.showOrginalName.Checked = False
                showOrigName = False
                ' Text neu berechnen 
                If Not IsNothing(selectedPlanShapes) Then
                    If selectedPlanShapes.Count = 1 Then
                        Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                        Me.elemName.Text = bestimmeElemText(tmpShape, showAbbrev.Checked, False)
                        ' wird im Formular immer lang dargestellt 
                        Me.elemDate.Text = bestimmeElemDateText(tmpShape, False)
                    End If
                End If



            ElseIf Not IsNothing(selectedPlanShapes) Then

                If selectedPlanShapes.Count = 1 Then
                    Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                    Me.elemName.Text = bestimmeElemText(tmpShape, showAbbrev.Checked, showOrginalName.Checked)
                    ' wird im Formular immer lang dargestellt 
                    Me.elemDate.Text = bestimmeElemDateText(tmpShape, False)
                End If

            End If
        Catch ex As Exception

        End Try


        dontFire = False

    End Sub

    Private Sub showOrginalName_CheckedChanged(sender As Object, e As EventArgs) Handles showOrginalName.CheckedChanged

        showOrigName = showOrginalName.Checked

        If dontFire Then
            Exit Sub
        End If

        Try
            If showOrginalName.Checked Then

                dontFire = True
                Me.showAbbrev.Checked = False
                showShortName = False

                ' Text neu berechnen 
                If Not IsNothing(selectedPlanShapes) Then
                    If selectedPlanShapes.Count = 1 Then
                        Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                        Me.elemName.Text = bestimmeElemText(tmpShape, False, True)
                    End If
                End If


            ElseIf Not IsNothing(selectedPlanShapes) Then

                If selectedPlanShapes.Count = 1 Then
                    Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                    Me.elemName.Text = bestimmeElemText(tmpShape, False, False)
                End If

            End If
        Catch ex As Exception

        End Try

        dontFire = False
    End Sub

    ''' <summary>
    ''' löscht die Text Annotation Strings
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deleteText_Click(sender As Object, e As EventArgs) Handles deleteText.Click
        Try
            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes

                    Call deleteAnnotationShape(tmpShape, pptAnnotationType.text)

                Next

                selectedPlanShapes.Select()

            End If

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' löscht das Text- bzw. Beschriftungs-TextElement des übergebenen Shapes  
    ''' </summary>
    ''' <param name="selectedPlanShape"></param>
    ''' <param name="descriptionType"></param>
    ''' <remarks></remarks>
    Private Sub deleteAnnotationShape(ByVal selectedPlanShape As PowerPoint.Shape, _
                                      ByVal descriptionType As Integer)

        Dim newShape As PowerPoint.Shape
        Dim textLeft As Double = selectedPlanShape.Left - 4
        Dim textTop As Double = selectedPlanShape.Top - 5
        Dim textwidth As Double = 5
        Dim textheight As Double = 5


        Dim shapeName As String = ""
        Dim ok As Boolean = False


        Try
            If Not IsNothing(descriptionType) Then
                If descriptionType >= 0 Then
                    shapeName = selectedPlanShape.Name & descriptionType.ToString
                    ok = True
                End If
            End If

        Catch ex As Exception
            ok = False
        End Try

        If Not ok Then
            Exit Sub
        End If

        Try
            newShape = currentSlide.Shapes(shapeName)
        Catch ex As Exception
            newShape = Nothing
        End Try


        If Not IsNothing(newShape) Then

            newShape.Delete()

        End If

    End Sub

    Private Sub positionTextButton_Click(sender As Object, e As EventArgs) Handles positionTextButton.Click
        Try
            If Not IsNothing(selectedPlanShapes) Then

                If selectedPlanShapes.Count = 1 Then
                    Dim isMilestone As Boolean = pptShapeIsMilestone(selectedPlanShapes.Item(1))
                    If isMilestone Then
                        positionIndexMT = positionIndexMT + 1
                        If positionIndexMT > 8 Then
                            positionIndexMT = 0
                        End If
                    Else
                        positionIndexPT = positionIndexPT + 1
                        If positionIndexPT > 8 Then
                            positionIndexPT = 0
                        End If
                    End If

                    Call Me.setDTPicture(isMilestone)
                End If

            End If
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' setzt die Bilder auf den Buttons zur Positionierung  
    ''' </summary>
    ''' <param name="isMilestone"></param>
    ''' <remarks></remarks>
    Public Sub setDTPicture(ByVal isMilestone As Boolean)

        Dim positionIndexT As Integer
        Dim positionIndexD As Integer

        If IsNothing(isMilestone) Then
            positionIndexD = -1
            positionIndexT = -1
        ElseIf isMilestone Then
            positionIndexD = Me.positionIndexMD
            positionIndexT = Me.positionIndexMT
        Else
            positionIndexD = Me.positionIndexPD
            positionIndexT = Me.positionIndexPT
        End If

        

        With Me
            Select Case positionIndexT
                Case -1
                    .positionTextButton.Image = Nothing
                Case pptPositionType.aboveCenter
                    .positionTextButton.Image = My.Resources.layout_north
                Case pptPositionType.aboveRight
                    .positionTextButton.Image = My.Resources.layout_northeast
                Case pptPositionType.centerRight
                    .positionTextButton.Image = My.Resources.layout_east
                Case pptPositionType.belowRight
                    .positionTextButton.Image = My.Resources.layout_southeast
                Case pptPositionType.belowCenter
                    .positionTextButton.Image = My.Resources.layout_south
                Case pptPositionType.belowLeft
                    .positionTextButton.Image = My.Resources.layout_southwest
                Case pptPositionType.centerLeft
                    .positionTextButton.Image = My.Resources.layout_west
                Case pptPositionType.aboveLeft
                    .positionTextButton.Image = My.Resources.layout_northwest
                Case pptPositionType.center
                    .positionTextButton.Image = My.Resources.layout_center
                Case Else
                    .positionTextButton.Image = My.Resources.layout_north
            End Select

            Select Case positionIndexD
                Case -1
                    .positionDateButton.Image = Nothing
                Case pptPositionType.aboveCenter
                    .positionDateButton.Image = My.Resources.layout_north
                Case pptPositionType.aboveRight
                    .positionDateButton.Image = My.Resources.layout_northeast
                Case pptPositionType.centerRight
                    .positionDateButton.Image = My.Resources.layout_east
                Case pptPositionType.belowRight
                    .positionDateButton.Image = My.Resources.layout_southeast
                Case pptPositionType.belowCenter
                    .positionDateButton.Image = My.Resources.layout_south
                Case pptPositionType.belowLeft
                    .positionDateButton.Image = My.Resources.layout_southwest
                Case pptPositionType.centerLeft
                    .positionDateButton.Image = My.Resources.layout_west
                Case pptPositionType.aboveLeft
                    .positionDateButton.Image = My.Resources.layout_northwest
                Case pptPositionType.center
                    .positionDateButton.Image = My.Resources.layout_center
                Case Else
                    .positionDateButton.Image = My.Resources.layout_north
            End Select
        End With


    End Sub

    Private Sub writeText_Click(sender As Object, e As EventArgs) Handles writeText.Click
        Try

            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    If pptShapeIsMilestone(tmpShape) Then
                        Call annotatePlanShape(tmpShape, pptAnnotationType.text, positionIndexMT)
                    Else
                        Call annotatePlanShape(tmpShape, pptAnnotationType.text, positionIndexPT)
                    End If

                Next

                selectedPlanShapes.Select()
            End If


        Catch ex As Exception

        End Try

    End Sub


    Private Sub deleteDate_Click(sender As Object, e As EventArgs) Handles deleteDate.Click
        Try

            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    Call deleteAnnotationShape(tmpShape, pptAnnotationType.datum)
                Next
                selectedPlanShapes.Select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub positionDateButton_Click(sender As Object, e As EventArgs) Handles positionDateButton.Click
        Try
            If Not IsNothing(selectedPlanShapes) Then
                If selectedPlanShapes.Count = 1 Then

                    Dim isMilestone As Boolean = pptShapeIsMilestone(selectedPlanShapes.Item(1))
                    If isMilestone Then
                        positionIndexMD = positionIndexMD + 1
                        If positionIndexMD > 8 Then
                            positionIndexMD = 0
                        End If
                    Else
                        positionIndexPD = positionIndexPD + 1
                        If positionIndexPD > 8 Then
                            positionIndexPD = 0
                        End If
                    End If

                    Call Me.setDTPicture(isMilestone)

                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub writeDate_Click(sender As Object, e As EventArgs) Handles writeDate.Click
        Try

            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    If pptShapeIsMilestone(tmpShape) Then
                        Call annotatePlanShape(tmpShape, pptAnnotationType.datum, positionIndexMD)
                    Elseif pptShapeIsPhase(tmpShape)
                        Call annotatePlanShape(tmpShape, pptAnnotationType.datum, positionIndexPD)
                    End If

                Next

                selectedPlanShapes.Select()
            End If

        Catch ex As Exception

        End Try


    End Sub

    Private Sub searchIcon_Click(sender As Object, e As EventArgs) Handles searchIcon.Click

        dontFire = True

        showSearchListBox = Not showSearchListBox

        If showSearchListBox Then
            Call searchBlockVisible(True)
            Call erstelleListbox()
        Else
            Call searchBlockVisible(False)
        End If

        dontFire = False

    End Sub


    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub writeAmpel_Click(sender As Object, e As EventArgs) Handles writeAmpel.Click
        Try

            Dim type As Integer
            If rdbMV.Checked Then
                type = pptAnnotationType.movedExplanation
            ElseIf rdbLU.Checked Then
                type = pptAnnotationType.lieferumfang
            Else
                type = pptAnnotationType.ampelText
            End If

            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes

                    If isRelevantShape(tmpShape) Then
                        If pptShapeIsMilestone(tmpShape) Then
                            Call annotatePlanShape(tmpShape, type, positionIndexMT)
                        Else
                            Call annotatePlanShape(tmpShape, type, positionIndexPT)
                        End If
                    End If


                Next

                selectedPlanShapes.Select()
            End If


        Catch ex As Exception

        End Try

    End Sub

    Private Sub deleteAmpel_Click(sender As Object, e As EventArgs) Handles deleteAmpel.Click
        Try

            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    Call deleteAnnotationShape(tmpShape, pptAnnotationType.ampelText)
                Next
                selectedPlanShapes.Select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureMarker_Click(sender As Object, e As EventArgs) Handles PictureMarker.Click
        CheckBxMarker.Checked = Not CheckBxMarker.Checked
    End Sub

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

    Private Sub rdbLU_CheckedChanged(sender As Object, e As EventArgs) Handles rdbLU.CheckedChanged

        If rdbLU.Checked = True Then

            Me.aLuTvText.Text = setALuTvText()

            ' prüfen , ob der AmpelBlock sichtbar ist ...
            If Me.aLuTvText.Visible Then
                ' alles ok 
            Else
                Call aLuTvBlockVisible(True)
            End If

            Call erstelleListbox()

        End If
    End Sub

    Private Sub rdbMV_CheckedChanged(sender As Object, e As EventArgs) Handles rdbMV.CheckedChanged
        If rdbMV.Checked = True Then

            Me.aLuTvText.Text = setALuTvText

            ' prüfen , ob der AmpelBlock sichtbar ist ...
            If Me.aLuTvText.Visible Then
                ' alles ok 
            Else
                Call aLuTvBlockVisible(True)
            End If

            Call erstelleListbox()

        End If
    End Sub

   
    ''' <summary>
    ''' im Falle: Termin-Veränderungen zeigen: alle in der Listbox markierten Elemente werden "auf Home-Position" geschickt ; wenn kein Element selektiert ist, dann alle 
    ''' im Fall eines selektierten Elements, das Home/Change Position hat: das oder die aktuell markierten Elemente werden zur Home-Position geschickt   
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSendToHome_Click(sender As Object, e As EventArgs) Handles btnSendToHome.Click

        Dim doItAll As Boolean = False
        If Not IsNothing(selectedPlanShapes) Then
            If selectedPlanShapes.Count > 0 Then
                ' alle selektierten Elemente zur Home-Position schicken 
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    If isRelevantShape(tmpShape) Then
                        Call sentToHomePosition(tmpShape)
                    End If
                Next
            Else
                doItAll = True
            End If
        Else
            doItAll = True
        End If

        If doItAll Then
            ' alle zur Home-Position schicken ...
            For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
                If isRelevantShape(tmpShape) Then
                    Call sentToHomePosition(tmpShape)
                End If
            Next
        End If

        ' jetzt ist Home nicht mehr notwendig ... 
        homeButtonRelevance = False

        btnSendToHome.Enabled = homeButtonRelevance
        btnSentToChange.Enabled = changedButtonRelevance

    End Sub

    ''' <summary>
    ''' im Falle: Termin-Veränderungen zeigen: alle in der Listbox markierten Elemente werden "auf Changed-Position" geschickt ; wenn kein Element selektiert ist, dann alle 
    ''' im Fall eines selektierten Elements, das Home/Change Position hat: das oder die aktuell markierten Elemente werden zur Changed-Position geschickt   
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSentToChange_Click(sender As Object, e As EventArgs) Handles btnSentToChange.Click

        Dim doItAll As Boolean = False
        If Not IsNothing(selectedPlanShapes) Then
            If selectedPlanShapes.Count > 0 Then
                ' alle selektierten Elemente zur Home-Position schicken 
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes

                    If isRelevantShape(tmpShape) Then
                        Call sentToChangedPosition(tmpShape.Name)
                    End If

                Next
            Else
                doItAll = True
                
            End If
        Else
            doItAll = True
        End If

        If doItAll Then
            ' alle zur Changed-Position schicken ...
            For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
                If isRelevantShape(tmpShape) Then
                    Call sentToChangedPosition(tmpShape.Name)
                End If
            Next

        End If

        changedButtonRelevance = False

        btnSentToChange.Enabled = changedButtonRelevance
        btnSendToHome.Enabled = homeButtonRelevance


    End Sub

    Private Sub aLuTvText_Enter(sender As Object, e As EventArgs) Handles aLuTvText.Enter

    End Sub

    Private Sub aLuTvText_TextChanged(sender As Object, e As EventArgs) Handles aLuTvText.TextChanged

    End Sub
End Class