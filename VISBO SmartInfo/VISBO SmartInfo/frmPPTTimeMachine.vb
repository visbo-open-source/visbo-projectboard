Public Class frmPPTTimeMachine
    Private currentTSIndex As Integer = -1
    Private timeStamps As SortedList(Of Date, Boolean) = Nothing

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
        btnForward.Enabled = False
        btnFastForward.Enabled = False
        btnEnd.Enabled = False
        btnBack.Enabled = False
        btnFastBack.Enabled = False
        btnStart.Enabled = False

        ' jetzt ggf wieder enabeln ...
        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                If currentTimestamp < timeStamps.Last.Key Then
                    btnForward.Enabled = True

                    If currentTimestamp.AddMonths(1) <= timeStamps.Last.Key Then
                        btnFastForward.Enabled = True
                    Else
                        btnFastForward.Enabled = False
                    End If

                    btnEnd.Enabled = True
                Else
                    btnForward.Enabled = False
                    btnFastForward.Enabled = False
                    btnEnd.Enabled = False
                End If


                If currentTimestamp > timeStamps.First.Key Then
                    btnBack.Enabled = True

                    If currentTimestamp.AddMonths(-1) >= timeStamps.First.Key Then
                        btnFastBack.Enabled = True
                    Else
                        btnFastBack.Enabled = False
                    End If

                    btnStart.Enabled = True
                Else
                    btnBack.Enabled = False
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

        currentTSIndex = -1
        ' gibt es ein Creation Date ?
        If smartSlideLists.creationDate > Date.MinValue Then
            txtboxCurrentDate.Text = currentTimestamp.ToShortDateString
        Else
            txtboxCurrentDate.Text = ""
        End If

        If noDBAccessInPPT Then
            Call MsgBox("no Database Access  ... action cancelled ...")
            MyBase.Close()
        Else
            ' gibt es überhaupt TimeStamps ? 
            timeStamps = smartSlideLists.getListOfTS


            If Not IsNothing(timeStamps) Then
                If timeStamps.Count >= 1 Then


                    txtboxCurrentDate.Enabled = True
                    lblMessage.Text = ""
                    Me.Text = "Time-Machine: " & timeStamps.First.Key.ToShortDateString & " - " & _
                        timeStamps.Last.Key.ToShortDateString & " (" & timeStamps.Count.ToString & ")"

                Else

                    txtboxCurrentDate.Enabled = False
                    txtboxCurrentDate.Text = ""
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
                currentTSIndex = timeStamps.Count - 1
                currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key

                txtboxCurrentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If


    End Sub

    Private Sub btnFastForward_Click(sender As Object, e As EventArgs) Handles btnFastForward.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                If currentTimestamp.AddMonths(1) < timeStamps.Last.Key Then
                    currentTimestamp = currentTimestamp.AddMonths(1)
                    ' jetzt den entsprechenden TSIndex auf -1 setzen, das heisst, er muss bestimmt werden ... 
                    currentTSIndex = -1
                Else
                    currentTSIndex = timeStamps.Count - 1
                    currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key
                End If

                txtboxCurrentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If

       

    End Sub

    Private Sub btnForward_Click(sender As Object, e As EventArgs) Handles btnForward.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim ix As Integer
                If currentTSIndex = -1 Then
                    ' was wäre der nächste ...
                    Dim found As Boolean = False
                    ix = 0
                    Do While ix <= timeStamps.Count - 1 And Not found
                        If DateDiff(DateInterval.Second, CDate(txtboxCurrentDate.Text), timeStamps.ElementAt(ix).Key) > 0 Then
                            found = True
                            currentTimestamp = timeStamps.ElementAt(ix).Key
                            currentTSIndex = ix
                        Else
                            ix = ix + 1
                        End If
                    Loop

                ElseIf currentTSIndex < timeStamps.Count - 1 Then
                    currentTSIndex = currentTSIndex + 1
                    currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key
                End If

                txtboxCurrentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If
        
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim ix As Integer
                If currentTSIndex = -1 Then
                    ' was wäre der nächste ...
                    Dim found As Boolean = False
                    ix = timeStamps.Count - 1
                    Do While ix >= 0 And Not found
                        If DateDiff(DateInterval.Second, CDate(txtboxCurrentDate.Text), timeStamps.ElementAt(ix).Key) < 0 Then
                            found = True
                            currentTimestamp = timeStamps.ElementAt(ix).Key
                            currentTSIndex = ix
                        Else
                            ix = ix - 1
                        End If
                    Loop

                ElseIf currentTSIndex >= 1 Then

                    currentTSIndex = currentTSIndex - 1
                    currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key

                End If

                txtboxCurrentDate.Text = currentTimestamp.ToString

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
                    ' jetzt den entsprechenden TSIndex auf -1 setzen, das heisst, er muss bestimmt werden ... 
                    currentTSIndex = -1
                Else
                    currentTSIndex = 0
                    currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key
                End If

                txtboxCurrentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)

            End If
        End If
        
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then
                currentTSIndex = 0
                currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key

                txtboxCurrentDate.Text = currentTimestamp.ToString

                Call moveAllShapes()

                Call showTSMessage(currentTimestamp)
            End If
        End If
        
    End Sub

    ' ''' <summary>
    ' ''' zeigt den Stand der Shapes zum angegebenen Zeitpunkt
    ' ''' </summary>
    ' ''' <param name="sender"></param>
    ' ''' <param name="e"></param>
    ' ''' <remarks></remarks>
    'Private Sub txtboxCurrentDate_Enter(sender As Object, e As EventArgs) Handles txtboxCurrentDate.Enter
    '    If Not IsNothing(timeStamps) Then
    '        If timeStamps.Count > 0 Then

    '            currentTimestamp = CDate(txtboxCurrentDate.Text)
    '            currentTSIndex = -1

    '            Call moveAllShapes()
    '            Call showTSMessage(currentTimestamp)

    '        End If
    '    End If
    'End Sub


    Private Sub btnForward_MouseHover(sender As Object, e As EventArgs) Handles btnForward.MouseHover

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim tmpDate As Date
                If currentTSIndex = -1 Then
                    ' was wäre der nächste ...
                    Dim found As Boolean = False
                    Dim ix As Integer = 0
                    Do While ix <= timeStamps.Count - 1 And Not found
                        If timeStamps.ElementAt(ix).Key > CDate(txtboxCurrentDate.Text) Then
                            found = True
                            tmpDate = timeStamps.ElementAt(ix).Key
                        Else
                            ix = ix + 1
                        End If
                    Loop

                ElseIf currentTSIndex < timeStamps.Count - 1 Then
                    tmpDate = timeStamps.ElementAt(currentTSIndex + 1).Key
                Else
                    tmpDate = timeStamps.Last.Key
                End If

                ToolTipTS.Show("Stand: " & _
                               tmpDate.ToString, btnForward, 2000)
            End If
        End If

        
    End Sub



    Private Sub btnFastForward_MouseHover(sender As Object, e As EventArgs) Handles btnFastForward.MouseHover

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then
                Dim tmpDate As Date = CDate(txtboxCurrentDate.Text).AddMonths(1)

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

   

    Private Sub btnBack_MouseHover(sender As Object, e As EventArgs) Handles btnBack.MouseHover


        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then
                Dim tmpDate As Date
                If Not IsNothing(timeStamps) Then

                    If timeStamps.Count > 0 Then
                        If currentTSIndex = -1 Then
                            ' was wäre der vorherige ...
                            Dim found As Boolean = False
                            Dim ix As Integer = timeStamps.Count - 1
                            Do While ix >= 0 And Not found
                                If timeStamps.ElementAt(ix).Key < CDate(txtboxCurrentDate.Text) Then
                                    found = True
                                    tmpDate = timeStamps.ElementAt(ix).Key
                                Else
                                    ix = ix - 1
                                End If
                            Loop

                        ElseIf currentTSIndex >= 1 Then
                            tmpDate = timeStamps.ElementAt(currentTSIndex - 1).Key
                        Else
                            tmpDate = timeStamps.First.Key
                        End If

                        ToolTipTS.Show("Stand: " & _
                                       tmpDate.ToString, btnBack, 2000)
                    End If
                End If
            End If
        End If

    End Sub

    

    Private Sub btnFastBack_MouseHover(sender As Object, e As EventArgs) Handles btnFastBack.MouseHover
        Dim tmpDate As Date = CDate(txtboxCurrentDate.Text).AddMonths(-1)


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

   
    Private Sub txtboxCurrentDate_ModifiedChanged(sender As Object, e As EventArgs) Handles txtboxCurrentDate.ModifiedChanged
        Call MsgBox("now")
    End Sub
End Class