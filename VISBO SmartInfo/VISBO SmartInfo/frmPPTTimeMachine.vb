Public Class frmPPTTimeMachine
    'Private currentTSIndex As Integer = -1
    Private timeStamps As SortedList(Of Date, Boolean) = Nothing
    Private noDateValueCheck As Boolean = True
    Private anzahlShapesOnSlide As Integer = currentSlide.Shapes.Count


    ''' <summary>
    ''' bewegt alle Shapes an 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub moveAllShapes()

        ' Progress-Bar anzeigen 
        ProgressBarNavigate.Value = 0
        ProgressBarNavigate.Visible = True

        Dim ix As Integer = 0

        ' alle Shapes zur Time-Stamp Position schicken ...
        ' in diffMvList wird gemerkt, um wieviel sich ein Shape verändert hat und ob überhaupt ...  
        Dim diffMvList As New SortedList(Of String, Double)
        Dim oldProgressValue = 0

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            ix = ix + 1

            If isRelevantShape(tmpShape) Then

                Call sendToTimeStampPosition(tmpShape, currentTimestamp, diffMvList)

            ElseIf isCommentShape(tmpShape) Then

                Call modifyComment(tmpShape, currentTimestamp)

            End If

            If CInt(10 * ix / anzahlShapesOnSlide) > oldProgressValue Then
                oldProgressValue = CInt(10 * ix / anzahlShapesOnSlide)
                ProgressBarNavigate.Value = oldProgressValue
            End If

        Next

        ' jetzt muss hier die Text- bzw Datums-Verschiebung laufen ... 
        ' die Diff-Werte stehen in der entsprechenden diffMvList

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

            Try

                If isAnnotationShape(tmpShape) Then

                    If tmpShape.Name.Substring(tmpShape.Name.Length - 1, 1) = pptAnnotationType.text Then

                        ' es handelt sich um den Text, also nur verschieben 
                        Dim refName As String = tmpShape.Name.Substring(0, tmpShape.Name.Length - 1)

                        If diffMvList.ContainsKey(refName) Then
                            Dim diff As Double = diffMvList.Item(refName)
                            With tmpShape
                                .Left = .Left + diff
                            End With
                        End If


                    ElseIf tmpShape.Name.Substring(tmpShape.Name.Length - 1, 1) = pptAnnotationType.datum Then

                        ' es handelt sich um das Datum, also verschieben und Text ändern 
                        Dim refName As String = tmpShape.Name.Substring(0, tmpShape.Name.Length - 1)
                        Dim refShape As PowerPoint.Shape = currentSlide.Shapes.Item(refName)
                        Dim tmpShort As Boolean = (tmpShape.TextFrame2.TextRange.Text.Length < 8)
                        Dim descriptionText As String = bestimmeElemDateText(refShape, tmpShort)

                        If diffMvList.ContainsKey(refName) Then
                            Dim diff As Double = diffMvList.Item(refName)
                            With tmpShape
                                .Left = .Left + diff
                                .TextFrame2.TextRange.Text = descriptionText
                            End With
                        End If

                    End If

                End If

            Catch ex As Exception
                Call MsgBox("Fehler : " & ex.Message)
            End Try

        Next

        Call buildSmartSlideLists()

        Call setBtnEnablements()

        ProgressBarNavigate.Visible = False
    End Sub

    Private Sub setBtnEnablements()


        ' alle Buttons erst mal auf enabled = false setzen  

        btnFastForward.Enabled = False
        btnEnd.Enabled = False

        btnFastBack.Enabled = False
        btnStart.Enabled = False

        ' jetzt ggf wieder enabeln ...
        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                If currentTimestamp < timeStamps.Last.Key Then


                    If currentTimestamp.AddMonths(1) <= timeStamps.Last.Key Then
                        btnFastForward.Enabled = True
                    Else
                        btnFastForward.Enabled = False
                    End If

                    btnEnd.Enabled = True
                Else

                    btnFastForward.Enabled = False
                    btnEnd.Enabled = False
                End If


                If currentTimestamp > timeStamps.First.Key Then


                    If currentTimestamp.AddMonths(-1) >= timeStamps.First.Key Then
                        btnFastBack.Enabled = True
                    Else
                        btnFastBack.Enabled = False
                    End If

                    btnStart.Enabled = True
                Else

                    btnFastBack.Enabled = False
                    btnStart.Enabled = False
                End If

            End If
        End If

        ' 
        If btnEnd.Enabled Then
            btnEnd.Focus()
        Else
            If btnStart.Enabled Then
                btnStart.Focus()
            End If
        End If


    End Sub

    Private Sub frmPPTTimeMachine_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    ''' <summary>
    ''' Laden der Time-Machine
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmPPTTimeMachine_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Progress-Bar visible ausschalten 
        ProgressBarNavigate.Visible = False
        ProgressBarNavigate.Value = 0

        'currentTSIndex = -1
        ' gibt es ein Creation Date ?
        If smartSlideLists.creationDate > Date.MinValue Then
            currentDate.Text = currentTimestamp.ToShortDateString
        Else
            currentDate.Text = ""
        End If

        If noDBAccessInPPT Then
            Call MsgBox("no Database Access  ... action cancelled ...")
            MyBase.Close()
        Else
            ' gibt es überhaupt TimeStamps ? 
            timeStamps = smartSlideLists.getListOfTS


            If Not IsNothing(timeStamps) Then
                If timeStamps.Count >= 1 Then


                    currentDate.Enabled = True
                    lblMessage.Text = ""
                    Me.Text = "Time-Machine: " & timeStamps.First.Key.ToShortDateString & " - " & _
                        timeStamps.Last.Key.ToShortDateString & " (" & timeStamps.Count.ToString & ")"

                Else

                    currentDate.Enabled = False
                    currentDate.Text = "" 'Date.Now.ToShortDateString
                    lblMessage.Text = "keine Einträge in der Datenbank vorhanden !"
                    Me.Text = "Time-Machine: "
                End If
            End If

            ' die beiden Buttons Home und ChangedPosition invisible setzen ..
            Call setBtnEnablements()

        End If

    End Sub

    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click

        If Not IsNothing(timeStamps) Then

            If timeStamps.Count > 0 Then

                currentTimestamp = timeStamps.Last.Key

                currentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If


    End Sub

    Private Sub btnFastForward_Click(sender As Object, e As EventArgs) Handles btnFastForward.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                If currentTimestamp.AddMonths(1) < Date.Now Then
                    currentTimestamp = currentTimestamp.AddMonths(1)
                Else
                    currentTimestamp = Date.Now
                End If

                currentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If

       

    End Sub


    Private Sub btnFastBack_Click(sender As Object, e As EventArgs) Handles btnFastBack.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                If currentTimestamp.AddMonths(-1) > timeStamps.First.Key Then
                    currentTimestamp = currentTimestamp.AddMonths(-1)
                Else
                    currentTimestamp = timeStamps.First.Key
                End If

                currentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If
        
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                currentTimestamp = timeStamps.First.Key

                currentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)
            End If
        End If
        
    End Sub


    Private Sub btnFastForward_MouseHover(sender As Object, e As EventArgs) Handles btnFastForward.MouseHover

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then
                Dim tmpDate As Date = CDate(currentDate.Text).AddMonths(1)

                If tmpDate < timeStamps.Last.Key Then
                    ' alles ok 
                    ToolTipTS.Show("Stand: " & _
                               tmpDate.ToString, btnFastForward, 2000)
                Else
                    tmpDate = timeStamps.Last.Key
                    ToolTipTS.Show("letzter Stand: " & _
                               tmpDate.ToString, btnFastForward, 2000)
                End If
            End If
        End If

        
    End Sub

    Private Sub btnEnd_MouseHover(sender As Object, e As EventArgs) Handles btnEnd.MouseHover
        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim tmpDate As Date = timeStamps.Last.Key
                ToolTipTS.Show("letzter Stand: " & _
                               tmpDate.ToString, btnEnd, 2000)

            End If
        End If

       
    End Sub

    

    Private Sub btnStart_MouseHover(sender As Object, e As EventArgs) Handles btnStart.MouseHover
        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim tmpDate As Date = timeStamps.First.Key
                ToolTipTS.Show("erster Stand: " & _
                               tmpDate.ToString, btnStart, 2000)

            End If
        End If
       
    End Sub


    Private Sub btnFastBack_MouseHover(sender As Object, e As EventArgs) Handles btnFastBack.MouseHover
        Dim tmpDate As Date = CDate(currentDate.Text).AddMonths(-1)


        If tmpDate > timeStamps.First.Key Then
            ' alles ok 
            ToolTipTS.Show("Stand: " & _
                       tmpDate.ToString, btnFastBack, 2000)
        Else
            tmpDate = timeStamps.First.Key
            ToolTipTS.Show("erster Stand: " & _
                       tmpDate.ToString, btnFastBack, 2000)
        End If

    End Sub

    Private Sub currentDate_GotFocus(sender As Object, e As EventArgs) Handles currentDate.GotFocus
        noDateValueCheck = False
    End Sub

    Private Sub currentDate_LostFocus(sender As Object, e As EventArgs) Handles currentDate.LostFocus
        noDateValueCheck = True
    End Sub


    Private Sub currentDate_ValueChanged(sender As Object, e As EventArgs) Handles currentDate.ValueChanged

        If noDateValueCheck Then
            Exit Sub
        Else

            If Not IsNothing(timeStamps) Then
                If timeStamps.Count > 0 Then
                    Dim eingabe As Date = CDate(currentDate.Text)
                    Try
                        ' ist es ein gültiges Datum ? 
                        If DateDiff(DateInterval.Day, eingabe, Date.Now) >= 0 And _
                            DateDiff(DateInterval.Day, eingabe, timeStamps.First.Key) <= 0 Then
                            ' es ist ein gültiges Datum ...
                            eingabe = CDate(currentDate.Text)
                        ElseIf DateDiff(DateInterval.Day, eingabe, Date.Now) < 0 Then
                            ' das Datum liegt in der Zukunft 

                            eingabe = Date.Now.ToShortDateString
                        ElseIf DateDiff(DateInterval.Day, eingabe, timeStamps.First.Key) > 0 Then
                            eingabe = timeStamps.First.Key.Date.AddHours(23).AddMinutes(50)
                        End If

                        currentTimestamp = eingabe
                        
                    Catch ex As Exception
                        Dim a As Integer = 0
                    End Try

                    noDateValueCheck = True
                    currentDate.Text = currentTimestamp.ToString
                    noDateValueCheck = False

                    Call moveAllShapes()

                    Call showTSMessage(currentTimestamp)

                End If
            End If
        End If
        


    End Sub
End Class