Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ProjectBoardBasic
Imports xlNS = Microsoft.Office.Interop.Excel
''' <summary>
''' das Form Info wird in variabler Größe angezeigt: mit / ohne Ampel-Block, mit /ohne Search-Block
''' es gibt zwei Methoden ampelblockVisibible und searchblockVisible, die die Elemente dann entsprechend positionieren und sichtbar machen 
''' </summary>
''' <remarks></remarks>
Public Class frmInfo

    Friend abkuerzung As String
    Friend showSearchListBox As Boolean = False

    ' tk , 16.5.
    ' Private Const deltaAmpel As Integer = 50
    Private Const deltaAmpel As Integer = 0
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
    ''' setzt im Fall englisch die Formular Texte auf englische Bezeichner 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub languageSettings()

        If englishLanguage Then
            With Me
                '.TabControl1.TabPages.Item(0).Text = "Information"
                '.TabControl1.TabPages.Item(1).Text = "Measure"
                '.showAbbrev.Text = "Abbreviation"
                '.lblAmpeln.Text = "Traffic-Light"
                '.lblAmpeln.Left = 432 - 20
                '.rdbLU.Text = "Deliverables"
                '.rdbMV.Text = "Changed Dates"
                '.rdbResources.Text = "Resources"
                '.rdbCosts.Text = "Cost"
                '.rdbAbbrev.Text = "Abbreviation"
                '.rdbBreadcrumb.Text = "full breadcrumb"
                '.rdbVerantwortlichkeiten.Text = "Responsibilities"
            End With
        End If

    End Sub

    Private Sub frmInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' sind irgendwelche Ampel-Farben gesetzt 
        Dim ix As Integer = 1

        ' sprach-Einstellung 
        Call languageSettings()

        formIsShown = True

        'PictureMarker.Visible = True
        'CheckBxMarker.Visible = True


        'With Me.shwGreenLight
        '    .Checked = showTrafficLights(1)
        '    .Visible = True
        'End With

        'With Me.shwYellowLight
        '    .Checked = showTrafficLights(2)
        '    .Visible = True
        'End With

        'With Me.shwRedLight
        '    .Checked = showTrafficLights(3)
        '    .Visible = True
        'End With

        'With Me.shwOhneLight
        '    .Checked = showTrafficLights(0)
        '    .Visible = True
        'End With

        'With Me.lblAmpeln
        '    .Visible = True
        'End With


        'CheckBxMarker.Checked = showMarker
        '' Zu Beginn ist Ampel-Text und Ampel-Erläuterung nicht sichtbar 
        'Call aLuTvBlockVisible(False)

        '' Zu Beginn ist die Searchbox nicht visible 
        'Call searchBlockVisible(False)

        dontFire = True

        showOrginalName.Visible = False
        showOrigName = False

        'If showBreadCrumbField = True Then
        '    fullBreadCrumb.Visible = True
        'Else
        '    fullBreadCrumb.Visible = False
        'End If


        showAbbrev.Checked = showShortName

        '' jetzt muss geprüft werden, ob GoToHome und GoToChangedPos enabled sind ... 
        'btnSentToChange.Enabled = changedButtonRelevance
        'btnSendToHome.Enabled = homeButtonRelevance

        dontFire = False

        '' ab jetzt sollen wieder die entsprechenden Event Routinen durchgeführt werden 
        'With Me.rdbName
        '    .Checked = True
        'End With

        ' wenn bereits was selektiert ist 
        If Not IsNothing(selectedPlanShapes) Then
            If selectedPlanShapes.Count = 1 Then
                Call aktualisiereInfoFrm(selectedPlanShapes(1))
            End If
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

    Private Sub showAbbrev_CheckedChanged(sender As Object, e As EventArgs)

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

    Private Sub showOrginalName_CheckedChanged(sender As Object, e As EventArgs)

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

    ''' <summary>
    ''' im Falle: Termin-Veränderungen zeigen: alle in der Listbox markierten Elemente werden "auf Home-Position" geschickt ; wenn kein Element selektiert ist, dann alle 
    ''' im Fall eines selektierten Elements, das Home/Change Position hat: das oder die aktuell markierten Elemente werden zur Home-Position geschickt   
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSendToHome_Click(sender As Object, e As EventArgs)

        Dim doItAll As Boolean = False
        If Not IsNothing(selectedPlanShapes) Then
            If selectedPlanShapes.Count > 0 Then
                ' alle selektierten Elemente zur Home-Position schicken 
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    If isRelevantMSPHShape(tmpShape) Then
                        Call sentToHomePosition(tmpShape.Name)
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
            Dim bigTodoList As New Collection
            For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
                bigTodoList.Add(tmpShape.Name)
            Next

            For Each tmpShpName As String In bigTodoList
                Try
                    Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                    If Not IsNothing(tmpShape) Then
                        If isRelevantMSPHShape(tmpShape) Then
                            Call sentToHomePosition(tmpShape.Name)
                        End If
                    End If
                Catch ex As Exception

                End Try
            Next


        End If

        ' jetzt ist Home nicht mehr notwendig ... 
        homeButtonRelevance = False

        'btnSendToHome.Enabled = homeButtonRelevance
        'btnSentToChange.Enabled = changedButtonRelevance

    End Sub

    ''' <summary>
    ''' im Falle: Termin-Veränderungen zeigen: alle in der Listbox markierten Elemente werden "auf Changed-Position" geschickt ; wenn kein Element selektiert ist, dann alle 
    ''' im Fall eines selektierten Elements, das Home/Change Position hat: das oder die aktuell markierten Elemente werden zur Changed-Position geschickt   
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSentToChange_Click(sender As Object, e As EventArgs)

        Dim doItAll As Boolean = False
        If Not IsNothing(selectedPlanShapes) Then
            If selectedPlanShapes.Count > 0 Then
                ' alle selektierten Elemente zur Home-Position schicken 
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes

                    If isRelevantMSPHShape(tmpShape) Then
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

            Dim bigTodoList As New Collection
            For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
                bigTodoList.Add(tmpShape.Name)
            Next

            For Each tmpShpName As String In bigTodoList
                Try
                    Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                    If Not IsNothing(tmpShape) Then
                        If isRelevantMSPHShape(tmpShape) Then
                            Call sentToChangedPosition(tmpShape.Name)
                        End If
                    End If
                Catch ex As Exception

                End Try
            Next

        End If

        changedButtonRelevance = False

        'btnSentToChange.Enabled = changedButtonRelevance
        'btnSendToHome.Enabled = homeButtonRelevance


    End Sub


    Private Sub deleteText_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "delete text annotation of element"
        Else
            tsMSG = "Löscht die Text Beschriftung des Elements"
        End If
        ToolTip1.Show(tsMSG, deleteText, 2000)
    End Sub

    Private Sub positionTextButton_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "relative position of text annotation to plan-element"
        Else
            tsMSG = "relative Position der Text-Beschriftung zum Plan-Element"
        End If
        ToolTip1.Show(tsMSG, positionTextButton, 2000)
    End Sub

    Private Sub writeText_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "annotate element with name"
        Else
            tsMSG = "erstellt die Text-Beschriftung des Elements"
        End If
        ToolTip1.Show(tsMSG, writeText, 2000)
    End Sub

    Private Sub deleteDate_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "delete date annotation of element"
        Else
            tsMSG = "Löscht die Datum-Beschriftung des Elements"
        End If
        ToolTip1.Show(tsMSG, deleteDate, 2000)
    End Sub

    Private Sub positionDateButton_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "relative position of date annotation to plan-element"
        Else
            tsMSG = "relative Position der Datum-Beschriftung zum Plan-Element"
        End If
        ToolTip1.Show(tsMSG, positionDateButton, 2000)
    End Sub

    Private Sub writeDate_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "annotate element with date"
        Else
            tsMSG = "erstellt die Datum-Beschriftung des Elements"
        End If
        ToolTip1.Show(tsMSG, writeDate, 2000)
    End Sub

   
    Private Sub showAbbrev_MouseHover(sender As Object, e As EventArgs)
        Dim tsMSG As String
        If englishLanguage Then
            tsMSG = "use abbreviation when annotating"
        Else
            tsMSG = "Verwende Kurzform zur Beschriftung"
        End If
        ToolTip1.Show(tsMSG, showAbbrev, 2000)
    End Sub

    Private Sub PictureMarker_MouseHover(sender As Object, e As EventArgs)
        'ToolTip1.Show("Element-Marker anzeigen", PictureMarker, 2000)
    End Sub

    Private Sub writeText_Click_1(sender As Object, e As EventArgs) Handles writeText.Click
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

    Private Sub writeDate_Click_1(sender As Object, e As EventArgs) Handles writeDate.Click
        Try

            If Not IsNothing(selectedPlanShapes) Then
                For Each tmpShape As PowerPoint.Shape In selectedPlanShapes
                    If pptShapeIsMilestone(tmpShape) Then
                        Call annotatePlanShape(tmpShape, pptAnnotationType.datum, positionIndexMD)
                    ElseIf pptShapeIsPhase(tmpShape) Then
                        Call annotatePlanShape(tmpShape, pptAnnotationType.datum, positionIndexPD)
                    End If

                Next

                selectedPlanShapes.Select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub positionTextButton_Click_1(sender As Object, e As EventArgs) Handles positionTextButton.Click
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

    Private Sub positionDateButton_Click_1(sender As Object, e As EventArgs) Handles positionDateButton.Click
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

    Private Sub deleteText_Click_1(sender As Object, e As EventArgs) Handles deleteText.Click
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

    Private Sub deleteDate_Click_1(sender As Object, e As EventArgs) Handles deleteDate.Click
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
End Class