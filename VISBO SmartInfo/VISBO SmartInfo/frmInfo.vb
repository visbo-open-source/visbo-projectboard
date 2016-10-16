Public Class frmInfo

    Friend abkuerzung As String
    Friend showSearchListBox As Boolean = False

    Friend Const fullHeight As Integer = 546
    Private Const smallHeight As Integer = 296

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
    End Sub

    Private Sub frmInfo_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        formIsShown = False
    End Sub

    Private Sub frmInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialisieren von Suchen 

        dontFire = True

        If showSearchListBox Then
            Me.Height = fullHeight
            filterText.Visible = True
            listboxNames.Visible = True
        Else
            Me.Height = smallHeight
            filterText.Visible = False
            listboxNames.Visible = False
        End If

        If showBreadCrumbField = True Then
            fullBreadCrumb.Visible = True
        Else
            fullBreadCrumb.Visible = False
        End If

        ' Anzeigen der Optionen oder nicht ...
        If extSearch = True Then
            rdbName.Visible = True
            rdbOriginalName.Visible = True
            rdbAbbrev.Visible = True
            rdbBreadcrumb.Visible = True
        Else
            rdbName.Visible = False
            rdbOriginalName.Visible = False
            rdbAbbrev.Visible = False
            rdbBreadcrumb.Visible = False
        End If

        ' sind irgendwelche Ampel-Farben gesetzt 
        Dim ix As Integer = 1

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


        dontFire = False

            ' ab jetzt sollen wieder die entsprechenden Event Routinen durchgeführt werden 
        With Me.rdbName
            .Checked = True
        End With

    End Sub

    Private Sub shwYellowLight_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub shwGreenLight_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    
    Private Sub rdbName_CheckedChanged(sender As Object, e As EventArgs) Handles rdbName.CheckedChanged
        ' dontFire true verhindert, dass die Aktion durchgeführt wird, das ist dann erforderlich wenn man explizit verhindern will, 
        ' dass ständig die Events getriggert werden 


        If rdbName.Checked = True Then

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
            Else
                rdbCode = pptInfoType.cName
            End If

            Dim nameCollection As Collection = smartSlideLists.getNCollection(colorCode, suchString, rdbCode)

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

            Call erstelleListbox()
            
        End If
    End Sub

    Private Sub rdbAbbrev_CheckedChanged(sender As Object, e As EventArgs) Handles rdbAbbrev.CheckedChanged
        If rdbAbbrev.Checked = True Then

            Call erstelleListbox()

        End If
    End Sub

    Private Sub rdbBreadcrumb_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBreadcrumb.CheckedChanged

        If rdbBreadcrumb.Checked = True Then

            Call erstelleListbox()

        End If

    End Sub

    Private Sub listboxNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listboxNames.SelectedIndexChanged

        ' es werden alle selektierten Namen als Shapes selektiert ....
        ' es können pro Name auch mehrere Shapes selektiert werden müssen 
        ' wenn Ampeln anzeigen an ist, dann werden auch die entsprechenden Ampel-Farben angezeigt ... 


        Dim nameArrayI() As String
        Dim nameArrayO() As String
        Dim anzSelected As Integer = listboxNames.SelectedItems.Count

        ReDim nameArrayI(anzSelected - 1)

        For i As Integer = 0 To anzSelected - 1
            nameArrayI(i) = CStr(listboxNames.SelectedItems.Item(i))
        Next

        Dim rdbCode As Integer

        If rdbName.Checked Then
            rdbCode = pptInfoType.cName
        ElseIf rdbOriginalName.Checked Then
            rdbCode = pptInfoType.oName
        ElseIf rdbAbbrev.Checked Then
            rdbCode = pptInfoType.sName
        ElseIf rdbBreadcrumb.Checked Then
            rdbCode = pptInfoType.bCrumb
        Else
            rdbCode = pptInfoType.cName
        End If

        Dim tmpCollection As Collection = smartSlideLists.getShapesNames(nameArrayI, rdbCode)

        anzSelected = tmpCollection.Count


        If anzSelected >= 1 Then
            ReDim nameArrayO(anzSelected - 1)

            For i As Integer = 0 To anzSelected - 1
                nameArrayO(i) = CStr(tmpCollection.Item(i + 1))
            Next

            Try
                selectedPlanShapes = currentSlide.Shapes.Range(nameArrayO)
                selectedPlanShapes.Select()
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
        Call erstelleListbox()
        Dim ampelColor As Integer = 0
        Call faerbeShapes(ampelColor, shwOhneLight.Checked)
    End Sub

    Private Sub shwGreenLight_CheckedChanged_1(sender As Object, e As EventArgs) Handles shwGreenLight.CheckedChanged
        Call erstelleListbox()
        Dim ampelColor As Integer = 1
        Call faerbeShapes(ampelColor, shwGreenLight.Checked)
    End Sub

    Private Sub shwYellowLight_CheckedChanged_1(sender As Object, e As EventArgs) Handles shwYellowLight.CheckedChanged
        Call erstelleListbox()
        Dim ampelColor As Integer = 2
        Call faerbeShapes(ampelColor, shwYellowLight.Checked)
    End Sub

    Private Sub shwRedLight_CheckedChanged(sender As Object, e As EventArgs) Handles shwRedLight.CheckedChanged
        Call erstelleListbox()
        Dim ampelColor As Integer = 3
        Call faerbeShapes(ampelColor, shwRedLight.Checked)
    End Sub

    Private Sub showAbbrev_CheckedChanged(sender As Object, e As EventArgs) Handles showAbbrev.CheckedChanged
        If dontFire Then
            Exit Sub
        End If

        Try
            If showAbbrev.Checked Then

                dontFire = True
                Me.showOrginalName.Checked = False

                ' Text neu berechnen 
                If Not IsNothing(selectedPlanShapes) Then
                    If selectedPlanShapes.Count = 1 Then
                        Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                        Me.elemName.Text = bestimmeElemText(tmpShape, True, False)
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

    Private Sub showOrginalName_CheckedChanged(sender As Object, e As EventArgs) Handles showOrginalName.CheckedChanged
        If dontFire Then
            Exit Sub
        End If

        Try
            If showOrginalName.Checked Then

                dontFire = True
                Me.showAbbrev.Checked = False

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

        If isMilestone Then
            positionIndexD = Me.positionIndexMD
            positionIndexT = Me.positionIndexMT
        Else
            positionIndexD = Me.positionIndexPD
            positionIndexT = Me.positionIndexPT
        End If

        With Me
            Select Case positionIndexT
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
                    .positionTextButton.Image = My.Resources.layout_horizontal
                Case Else
                    .positionTextButton.Image = My.Resources.layout_north
            End Select

            Select Case positionIndexD
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
                    .positionDateButton.Image = My.Resources.layout_horizontal
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
    ''' <summary>
    ''' fügt in der Powerpoint an das selektierte Plan-Element Lang-Name, Original-Name, Kurz-Name bzw Datum an 
    ''' wenn das Element bereits existiert, so wird es mit dem betreffenden Text beschriftet   
    ''' globale Variable, die im Zugriff sind: 
    ''' currentSlide: die aktuelle PPT-Slide
    ''' selectedplanShape: das aktuell selektierte Plan-Shape 
    ''' </summary>
    ''' <param name="descriptionType"></param>
    ''' <param name="positionIndex"></param>
    ''' <remarks></remarks>
    Private Sub annotatePlanShape(ByVal selectedPlanShape As PowerPoint.Shape, _
                                  ByVal descriptionType As Integer, ByVal positionIndex As Integer)

        Dim newShape As PowerPoint.Shape
        Dim textLeft As Double = selectedPlanShape.Left - 4
        Dim textTop As Double = selectedPlanShape.Top - 5
        Dim textwidth As Double = 5
        Dim textheight As Double = 5
        Dim normalFarbe As Integer = RGB(10, 10, 10)

        Dim descriptionText As String = ""

        Dim shapeName As String = ""
        Dim ok As Boolean = False

        ' handelt es sich um den Lang-/Kurz-Namen oder um das Datum ? 

        If descriptionType = pptAnnotationType.text Then
            descriptionText = bestimmeElemText(selectedPlanShape, Me.showAbbrev.Checked, Me.showOrginalName.Checked)
        ElseIf descriptionType = pptAnnotationType.datum Then
            descriptionText = bestimmeElemDateText(selectedPlanShape)
        End If

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


        If IsNothing(newShape) Then

            newShape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, _
                                      textLeft, textTop, 50, textheight)
            With newShape
                .TextFrame2.TextRange.Text = descriptionText
                .TextFrame2.TextRange.Font.Size = CDbl(schriftGroesse)
                .TextFrame2.MarginBottom = 0
                .TextFrame2.MarginLeft = 0
                .TextFrame2.MarginRight = 0
                .TextFrame2.MarginTop = 0
                .Name = shapeName
                .TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
            End With

        Else
            With newShape
                .TextFrame2.TextRange.Text = descriptionText
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = normalFarbe
            End With
        End If


        ' jetzt wird das TextShape noch positioniert - in Abhängigkeit vom Position Index 

        Select Case positionIndex

            Case pptPositionType.center

                If newShape.Width > 1.5 * selectedPlanShape.Width Then
                    ' keine Farbänderung 
                Else
                    ' wenn die Beschriftung von der Ausdehnung kleiner als die Phase/der Meilenstein ist
                    newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                End If
                textLeft = selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width)
                textTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

            Case pptPositionType.aboveCenter

                textLeft = selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width)
                textTop = selectedPlanShape.Top - newShape.Height

            Case pptPositionType.aboveRight

                If newShape.Width > selectedPlanShape.Width Then
                    textLeft = selectedPlanShape.Left
                Else
                    textLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                End If

                textTop = selectedPlanShape.Top - newShape.Height

            Case pptPositionType.centerRight

                If newShape.Width > selectedPlanShape.Width Then
                    textLeft = selectedPlanShape.Left
                Else
                    textLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If

                End If

                textTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

            Case pptPositionType.belowRight

                If newShape.Width > selectedPlanShape.Width Then
                    textLeft = selectedPlanShape.Left
                Else
                    textLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                End If

                textTop = selectedPlanShape.Top + selectedPlanShape.Height

            Case pptPositionType.belowCenter
                textLeft = selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width)
                textTop = selectedPlanShape.Top + selectedPlanShape.Height

            Case pptPositionType.belowLeft

                If newShape.Width > selectedPlanShape.Width Then
                    textLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                Else
                    textLeft = selectedPlanShape.Left
                    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                End If

                textTop = selectedPlanShape.Top + selectedPlanShape.Height

            Case pptPositionType.centerLeft
                If newShape.Width > selectedPlanShape.Width Then
                    textLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                Else
                    textLeft = selectedPlanShape.Left
                    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                End If
                textTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

            Case pptPositionType.aboveLeft
                If newShape.Width > selectedPlanShape.Width Then
                    textLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                Else
                    textLeft = selectedPlanShape.Left
                    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                End If
                textTop = selectedPlanShape.Top - newShape.Height

            Case Else
                textLeft = selectedPlanShape.Left - 5
                textTop = selectedPlanShape.Top - 10
        End Select

        ' jetzt die Position zuweisen

        With newShape
            .Top = textTop
            .Left = textLeft
        End With


        'currentSlide.Shapes(newShape.Name).Select()


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
                    Else
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
            Me.Height = fullHeight
            filterText.Visible = True
            listboxNames.Visible = True

            Call erstelleListbox()
        Else
            Me.Height = smallHeight
            filterText.Visible = False
            listboxNames.Visible = False
        End If

        dontFire = False

    End Sub

    
End Class