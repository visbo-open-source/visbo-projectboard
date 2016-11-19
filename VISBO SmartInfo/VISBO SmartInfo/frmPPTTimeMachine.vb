Public Class frmPPTTimeMachine
    Private currentTSIndex As Integer = -1
    Private currentTimestamp As Date = Date.MinValue
    Private timeStamps As SortedList(Of Date, Boolean) = Nothing

    Private anzahlShapesOnSlide As Integer = currentSlide.Shapes.Count


    ''' <summary>
    ''' bewegt alle Shapes an 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub moveAllShapes(Optional ByVal toHomePosition As Boolean = False)

        If toHomePosition Then
            currentTSIndex = -1
            currentTimestamp = smartSlideLists.creationDate
            txtboxCurrentDate.Text = "Home-Position"
        Else
            txtboxCurrentDate.Text = currentTimestamp.ToString
        End If


        ' Progress-Bar anzeigen 
        ProgressBarNavigate.Value = 0
        ProgressBarNavigate.Visible = True

        Dim ix As Integer = 0

        ' alle Shapes zur Time-Stamp Position schicken ...
        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            ix = ix + 1

            If isRelevantShape(tmpShape) Then
                If Not toHomePosition Then
                    Call sendToTimeStampPosition(tmpShape, currentTimestamp)
                Else
                    Call sentToHomePosition(tmpShape)
                End If
            End If

            ProgressBarNavigate.Value = CInt(10 * ix / anzahlShapesOnSlide)

        Next

        ' sind die Home / Change-Buttons visible, enabled 
        btnHome.Visible = True
        If toHomePosition Then
            homeButtonRelevance = False
            btnHome.Enabled = homeButtonRelevance
        Else
            btnHome.Enabled = homeButtonRelevance
        End If

        btnChangedPosition.Visible = True
        If Not toHomePosition Then
            changedButtonRelevance = False
            btnChangedPosition.Enabled = changedButtonRelevance
        Else
            btnChangedPosition.Enabled = changedButtonRelevance
        End If

        ' was ist mit den Navigate buttons ... 

        If currentTimestamp < timeStamps.Last.Key Then
            btnForward.Enabled = True
            btnFastForward.Enabled = True
            btnEnd.Enabled = True
        Else
            btnForward.Enabled = False
            btnFastForward.Enabled = False
            btnEnd.Enabled = False
        End If


        If currentTimestamp > timeStamps.First.Key Then
            btnBack.Enabled = True
            btnFastBack.Enabled = True
            btnStart.Enabled = True
        Else
            btnBack.Enabled = False
            btnFastBack.Enabled = False
            btnStart.Enabled = False
        End If

        ProgressBarNavigate.Visible = False
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
            currentTimestamp = smartSlideLists.creationDate
            txtboxCurrentDate.Text = currentTimestamp.ToShortDateString
        End If

        If noDBAccessInPPT Then
            Call MsgBox("kein Datenbank Zugriff ... Abbruch ...")
            MyBase.Close()
        Else
            ' die beiden Buttons Home und ChangedPosition invisible setzen ..
            btnChangedPosition.Visible = False
            btnHome.Visible = False

            ' die Navigation de-aktivieren 

            ' gibt es überhaupt TimeStamps ? 
            timeStamps = smartSlideLists.getListOfTS

            If Not IsNothing(timeStamps) Then
                If timeStamps.Count >= 1 Then

                    ' die Navigation enablen  
                    btnBack.Enabled = False
                    btnFastBack.Enabled = False
                    btnStart.Enabled = True
                    btnForward.Enabled = False
                    btnFastForward.Enabled = False
                    btnEnd.Enabled = True
                    txtboxCurrentDate.Enabled = True
                    txtboxCurrentDate.Text = ""
                    lblMessage.Text = ""
                    Me.Text = "Time-Machine: " & timeStamps.First.Key.ToShortDateString & " - " & _
                        timeStamps.Last.Key.ToShortDateString & " (" & timeStamps.Count.ToString & ")"

                Else

                    ' die Navigation dis-abeln
                    btnBack.Enabled = False
                    btnFastBack.Enabled = False
                    btnStart.Enabled = False
                    btnForward.Enabled = False
                    btnFastForward.Enabled = False
                    btnEnd.Enabled = False
                    txtboxCurrentDate.Enabled = False
                    txtboxCurrentDate.Text = ""
                    lblMessage.Text = "keine Einträge in der Datenbank vorhanden !"

                End If
            End If



        End If

    End Sub

    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click

        If Not IsNothing(timeStamps) Then

            If timeStamps.Count > 0 Then
                currentTSIndex = timeStamps.Count - 1
                currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key

                Call moveAllShapes()

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

                Call moveAllShapes()

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
                        If timeStamps.ElementAt(ix).Key > CDate(txtboxCurrentDate.Text) Then
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

                Call moveAllShapes()

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
                        If timeStamps.ElementAt(ix).Key < CDate(txtboxCurrentDate.Text) Then
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

                Call moveAllShapes()
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

                Call moveAllShapes()

            End If
        End If
        
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then
                currentTSIndex = 0
                currentTimestamp = timeStamps.ElementAt(currentTSIndex).Key

                Call moveAllShapes()
            End If
        End If
        
    End Sub

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

    Private Sub btnHome_Click(sender As Object, e As EventArgs) Handles btnHome.Click

        ' true bedeutet: move to Home Position
        Call moveAllShapes(True)

    End Sub

    Private Sub btnChangedPosition_Click(sender As Object, e As EventArgs) Handles btnChangedPosition.Click

        ' false bedeutet: move to Changed Position 
        Call moveAllShapes(False)

    End Sub
End Class