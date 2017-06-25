Public Class frmPPTTimeMachine
    'Private currentTSIndex As Integer = -1
    Private timeStamps As SortedList(Of Date, Boolean) = Nothing
    Private timeStampsIndex As Integer = -1
    Private noDateValueCheck As Boolean = True
    Private anzahlShapesOnSlide As Integer = currentSlide.Shapes.Count



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

                    If smartSlideLists.countProjects = 1 Then
                        
                        btnEnd.Enabled = True
                        btnFastForward.Enabled = True

                    Else
                        If currentTimestamp.AddMonths(1) <= timeStamps.Last.Key Then
                            btnFastForward.Enabled = True
                        Else
                            btnFastForward.Enabled = False
                        End If
                        btnEnd.Enabled = True
                    End If
                    

                Else
                    btnFastForward.Enabled = False
                    btnEnd.Enabled = False
                End If


                If currentTimestamp > timeStamps.First.Key Then

                    If smartSlideLists.countProjects = 1 Then
                        btnStart.Enabled = True
                        btnFastBack.Enabled = True

                    Else
                        If currentTimestamp.AddMonths(-1) >= timeStamps.First.Key Then
                            btnFastBack.Enabled = True
                        Else
                            btnFastBack.Enabled = False
                        End If

                        btnStart.Enabled = True

                    End If
                    
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

                    ' bestimme hier aufgrund des Datums den timestampsIndex
                    If timeStamps.Count > 0 Then
                        If smartSlideLists.countProjects = 1 Then
                            ' nimm das Datum, das in der sortierten Liste unmittelbar davor liegt 
                            Dim ix As Integer = timeStamps.Count - 1
                            Dim found As Boolean = False
                            Do While ix >= 0 And Not found
                                If currentTimestamp >= timeStamps.ElementAt(ix).Key Then
                                    found = True
                                Else
                                    ix = ix - 1
                                End If
                            Loop

                            If found Then
                                timeStampsIndex = ix
                            End If
                        Else
                            ' ist ja schon gesetzt 
                        End If
                    End If


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

                timeStampsIndex = timeStamps.Count - 1

                If timeStamps.Last.Key <> currentTimestamp Then

                    previousVariantName = currentVariantname
                    previousTimeStamp = currentTimestamp
                    currentTimestamp = timeStamps.Last.Key

                    currentDate.Text = currentTimestamp.ToString

                    Call moveAllShapes()

                    Call setBtnEnablements()

                    Call setCurrentTimeStampInSlide(currentTimestamp)

                    If thereIsNoVersionFieldOnSlide Then
                        Call showTSMessage(currentTimestamp)
                    End If

                End If


            End If
        End If


    End Sub

    Private Sub btnFastForward_Click(sender As Object, e As EventArgs) Handles btnFastForward.Click
        Dim newDate As Date
        Dim found As Boolean = False
        Dim weitermachen As Boolean = False


        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                timeStampsIndex = timeStampsIndex + 1
                If timeStampsIndex > timeStamps.Count - 1 Then
                    timeStampsIndex = timeStamps.Count - 1
                End If

                If smartSlideLists.countProjects = 1 Then
                    newDate = timeStamps.ElementAt(timeStampsIndex).Key
                Else
                    If currentTimestamp.AddMonths(-1) > timeStamps.First.Key Then
                        newDate = currentTimestamp.AddMonths(-1)
                    Else
                        newDate = timeStamps.First.Key
                    End If
                End If


                If newDate <> currentTimestamp Then
                    previousVariantName = currentVariantname
                    previousTimeStamp = currentTimestamp
                    currentTimestamp = newDate
                    currentDate.Text = currentTimestamp.ToString

                    Call moveAllShapes()

                    Call setBtnEnablements()

                    Call setCurrentTimestampInSlide(currentTimestamp)

                    If thereIsNoVersionFieldOnSlide Then
                        Call showTSMessage(currentTimestamp)
                    End If

                End If


            End If
        End If


    End Sub


    Private Sub btnFastBack_Click(sender As Object, e As EventArgs) Handles btnFastBack.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim newDate As Date
                timeStampsIndex = timeStampsIndex - 1
                If timeStampsIndex < 0 Then
                    timeStampsIndex = 0
                End If

                If smartSlideLists.countProjects = 1 Then
                    newDate = timeStamps.ElementAt(timeStampsIndex).Key
                Else
                    If currentTimestamp.AddMonths(-1) > timeStamps.First.Key Then
                        newDate = currentTimestamp.AddMonths(-1)
                    Else
                        newDate = timeStamps.First.Key
                    End If
                End If
                

                If newDate <> currentTimestamp Then
                    previousVariantName = currentVariantname
                    previousTimeStamp = currentTimestamp
                    currentTimestamp = newDate
                    currentDate.Text = currentTimestamp.ToString

                    Call moveAllShapes()

                    Call setBtnEnablements()

                    Call setCurrentTimestampInSlide(currentTimestamp)

                    If thereIsNoVersionFieldOnSlide Then
                        Call showTSMessage(currentTimestamp)
                    End If

                End If
                

            End If
        End If
        
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                timeStampsIndex = 0
                If timeStamps.First.Key <> currentTimestamp Then
                    previousVariantName = currentVariantname
                    previousTimeStamp = currentTimestamp
                    currentTimestamp = timeStamps.First.Key

                    currentDate.Text = currentTimestamp.ToString

                    Call moveAllShapes()

                    Call setBtnEnablements()

                    Call setCurrentTimestampInSlide(currentTimestamp)

                    If thereIsNoVersionFieldOnSlide Then
                        Call showTSMessage(currentTimestamp)
                    End If

                End If
                
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

                    Dim eingabe As Date = CDate(currentDate.Text).Date.AddHours(23).AddMinutes(59)
                    Try
                        ' ist es ein gültiges Datum ? 
                        If DateDiff(DateInterval.Day, eingabe, Date.Now) >= 0 And _
                            DateDiff(DateInterval.Day, eingabe, timeStamps.First.Key) <= 0 Then
                            ' es ist ein gültiges Datum ...

                            If smartSlideLists.countProjects = 1 Then
                                ' nimm das Datum, das in der sortierten Liste unmittelbar davor liegt 
                                Dim ix As Integer = timeStamps.Count - 1
                                Dim found As Boolean = False
                                Do While ix >= 0 And Not found
                                    If eingabe >= timeStamps.ElementAt(ix).Key Then
                                        found = True
                                    Else
                                        ix = ix - 1
                                    End If
                                Loop

                                If found Then
                                    timeStampsIndex = ix
                                End If
                            Else
                                ' ist ja schon gesetzt 
                            End If

                        ElseIf DateDiff(DateInterval.Day, eingabe, Date.Now) < 0 Then
                            ' das Datum liegt in der Zukunft 

                            eingabe = timeStamps.Last.Key.AddMinutes(1)
                            timeStampsIndex = timeStamps.Count - 1

                        ElseIf DateDiff(DateInterval.Day, eingabe, timeStamps.First.Key) > 0 Then
                            eingabe = timeStamps.First.Key.AddMinutes(1)
                            timeStampsIndex = 0

                        End If



                    Catch ex As Exception
                        Dim a As Integer = 0
                    End Try

                    If eingabe <> currentTimestamp Then

                        previousVariantName = currentVariantname
                        previousTimeStamp = currentTimestamp
                        currentTimestamp = eingabe
                        noDateValueCheck = True
                        currentDate.Text = currentTimestamp.ToString
                        noDateValueCheck = False

                        Call moveAllShapes()
                        Call setBtnEnablements()

                        Call setCurrentTimestampInSlide(currentTimestamp)

                        If thereIsNoVersionFieldOnSlide Then
                            Call showTSMessage(currentTimestamp)
                        End If

                    End If


                End If
            End If
        End If
        


    End Sub
End Class