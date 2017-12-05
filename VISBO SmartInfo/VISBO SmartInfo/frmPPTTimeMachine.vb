Public Class frmPPTTimeMachine
    'Private currentTSIndex As Integer = -1
    Private timeStamps As SortedList(Of Date, Boolean) = Nothing
    Private timeStampsIndex As Integer = -1
    Private anzahlShapesOnSlide As Integer = currentSlide.Shapes.Count


    Private Enum ptNavigationButtons
        letzter = 0
        erster = 1
        nachher = 2
        vorher = 3
        individual = 4
    End Enum

    Private Sub setBtnEnablements()

        ' alle Buttons erst mal auf enabled = false setzen  

        btnFastForward.Enabled = False
        btnEnd.Enabled = False
        btnEnd2.Enabled = False

        btnFastBack.Enabled = False
        btnStart.Enabled = False

        ' jetzt ggf wieder enabeln ...
        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                If currentTimestamp < timeStamps.Last.Key Then

                    ' Änderung tk 13.8.17 , btnEnd und btnFastforward immer enablen ... 
                    btnEnd.Enabled = True
                    btnEnd2.Enabled = True
                    btnFastForward.Enabled = True

                    ''If smartSlideLists.countProjects = 1 Then

                    ''    btnEnd.Enabled = True
                    ''    btnFastForward.Enabled = True

                    ''Else
                    ''    'If currentTimestamp.AddMonths(1) <= timeStamps.Last.Key Then
                    ''    '    btnFastForward.Enabled = True
                    ''    'Else
                    ''    '    btnFastForward.Enabled = False
                    ''    'End If
                    ''    btnFastForward.Enabled = True
                    ''    btnEnd.Enabled = True
                    ''End If


                Else
                    btnEnd.Enabled = False
                    btnEnd2.Enabled = False
                    btnFastForward.Enabled = False
                End If


                If currentTimestamp > timeStamps.First.Key Then

                    ' Änderung tk 13.8.17 , btnStart und btnFastBack immer enablen ... 
                    btnStart.Enabled = True
                    btnFastBack.Enabled = True

                    ''If smartSlideLists.countProjects = 1 Then
                    ''    btnStart.Enabled = True
                    ''    btnFastBack.Enabled = True

                    ''Else
                    ''    If currentTimestamp.AddMonths(-1) >= timeStamps.First.Key Then
                    ''        btnFastBack.Enabled = True
                    ''    Else
                    ''        btnFastBack.Enabled = False
                    ''    End If

                    ''    btnStart.Enabled = True

                    ''End If

                Else

                    btnStart.Enabled = False
                    btnFastBack.Enabled = False

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

        ' zu Beginn die Checkbox Changelist nicht anzeigen ...
        showChangeList.Visible = False

        ' die MArker, falls welche sichtbar sind , wegmachen ... 
        Call deleteMarkerShapes()

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
    ''' <summary>
    ''' führt die Button Action der Time-Machine aus 
    ''' </summary>
    ''' <param name="newdate"></param>
    ''' <remarks></remarks>
    Private Sub performBtnAction(ByVal newdate As Date)


        If newdate <> currentTimestamp Then


            Me.UseWaitCursor = True
            ' clear changelist 
            'Call changeListe.clearChangeList()

            previousVariantName = currentVariantname
            previousTimeStamp = currentTimestamp
            currentTimestamp = newdate

            currentDate.Text = currentTimestamp.ToString

            Call moveAllShapes()

            Call setBtnEnablements()

            Call setCurrentTimestampInSlide(currentTimestamp)

            If thereIsNoVersionFieldOnSlide Then
                Call showTSMessage(currentTimestamp)
            End If

            ' jetzt prüfen, ob es Veränderungen im PPT gab, aktuell beschränkt auf Meilensteine und Phasen ..
            If showChangeList.Checked = True Then
                ' das Formular aufschalten 
                If IsNothing(changeFrm) Then
                    changeFrm = New frmChanges
                    changeFrm.Show()
                Else
                    changeFrm.neuAufbau()
                End If
            End If

            Me.UseWaitCursor = False

        End If

    End Sub
    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click

        If Not IsNothing(timeStamps) Then

            If timeStamps.Count > 0 Then

                Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.letzter)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

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

                newDate = getNextNavigationDate(ptNavigationButtons.nachher)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

                End If


            End If
        End If


    End Sub


    Private Sub btnFastBack_Click(sender As Object, e As EventArgs) Handles btnFastBack.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.vorher)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

                End If


            End If
        End If

    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click

        If Not IsNothing(timeStamps) Then
            If timeStamps.Count > 0 Then

                Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.erster)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

                End If

            End If
        End If

    End Sub


    Private Sub btnFastForward_MouseHover(sender As Object, e As EventArgs) Handles btnFastForward.MouseHover

        Dim tmpDate As Date = getNextNavigationDate(ptNavigationButtons.nachher, True)
        ToolTipTS.Show(tmpDate.ToString, btnFastForward, 2000)


    End Sub

    Private Sub btnEnd_MouseHover(sender As Object, e As EventArgs) Handles btnEnd.MouseHover

        Dim tmpDate As Date = getNextNavigationDate(ptNavigationButtons.letzter, True)
        ToolTipTS.Show(tmpDate.ToString, btnEnd, 2000)

    End Sub



    Private Sub btnStart_MouseHover(sender As Object, e As EventArgs) Handles btnStart.MouseHover

        Dim tmpDate As Date = getNextNavigationDate(ptNavigationButtons.erster, True)
        ToolTipTS.Show(tmpDate.ToString, btnStart, 2000)

    End Sub


    Private Sub btnFastBack_MouseHover(sender As Object, e As EventArgs) Handles btnFastBack.MouseHover

        Dim tmpDate As Date = getNextNavigationDate(ptNavigationButtons.vorher, True)
        ToolTipTS.Show(tmpDate.ToString, btnFastBack, 2000)

    End Sub



    Private Sub updateWithNewDate()
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

                    '' jetzt die Checkbox anzeigen ... 
                    Me.showChangeList.Visible = True

                    ' clear changelist 
                    Call changeListe.clearChangeList()

                    previousVariantName = currentVariantname
                    previousTimeStamp = currentTimestamp
                    currentTimestamp = eingabe

                    currentDate.Text = currentTimestamp.ToString

                    Call moveAllShapes()
                    Call setBtnEnablements()

                    Call setCurrentTimestampInSlide(currentTimestamp)

                    If thereIsNoVersionFieldOnSlide Then
                        Call showTSMessage(currentTimestamp)
                    End If

                    ' jetzt prüfen, ob es Veränderungen im PPT gab, aktuell beschränkt auf Meilensteine und Phasen ..
                    If showChangeList.Checked = True Then
                        ' das Formular aufschalten 
                        If IsNothing(changeFrm) Then
                            changeFrm = New frmChanges
                            changeFrm.Show()
                        Else
                            changeFrm.neuAufbau()
                        End If
                    End If

                End If


            End If
        End If
    End Sub

    ''' <summary>
    ''' gibt das Datum zurück, das eingestellt wird, wenn der Button gedrückt wird ... 
    ''' wenn irgendwas schief get 
    ''' </summary>
    ''' <param name="kennung"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getNextNavigationDate(ByVal kennung As Integer, Optional ByVal justForInformation As Boolean = False) As Date
        Dim tmpDate As Date = Date.Now
        Dim tmpIndex As Integer = timeStampsIndex

        Select Case kennung
            Case ptNavigationButtons.nachher


                If timeStamps.Count > 0 Then
                    tmpIndex = tmpIndex + 1

                    If tmpIndex > timeStamps.Count - 1 Then
                        tmpIndex = timeStamps.Count - 1
                    End If

                    If smartSlideLists.countProjects = 1 Then
                        tmpDate = timeStamps.ElementAt(tmpIndex).Key
                    Else
                        If currentTimestamp.AddMonths(1) < timeStamps.Last.Key Then
                            tmpDate = currentTimestamp.AddMonths(1)
                        Else
                            tmpDate = timeStamps.Last.Key
                        End If
                    End If
                End If

            Case ptNavigationButtons.vorher

                If timeStamps.Count > 0 Then
                    tmpIndex = tmpIndex - 1

                    If tmpIndex < 0 Then
                        tmpIndex = 0
                    End If

                    If smartSlideLists.countProjects = 1 Then
                        tmpDate = timeStamps.ElementAt(tmpIndex).Key
                    Else
                        If currentTimestamp.AddMonths(-1) > timeStamps.First.Key Then
                            tmpDate = currentTimestamp.AddMonths(-1)
                        Else
                            tmpDate = timeStamps.First.Key
                        End If
                    End If
                End If


            Case ptNavigationButtons.erster

                If timeStamps.Count > 0 Then
                    tmpIndex = 0
                    tmpDate = timeStamps.First.Key
                End If

            Case ptNavigationButtons.letzter


                If timeStamps.Count > 0 Then
                    tmpIndex = timeStamps.Count - 1
                    tmpDate = timeStamps.Last.Key
                End If


        End Select

        If Not justForInformation Then
            timeStampsIndex = tmpIndex
        End If

        getNextNavigationDate = tmpDate
    End Function

    Private Sub showChangeList_CheckedChanged(sender As Object, e As EventArgs) Handles showChangeList.CheckedChanged

        If showChangeList.Checked Then
            ' das Formular aufschalten 
            If IsNothing(changeFrm) Then
                changeFrm = New frmChanges
                changeFrm.Show()
            Else
                changeFrm.neuAufbau()
            End If

        Else
            Try
                If Not IsNothing(changeFrm) Then
                    changeFrm.Close()
                    changeFrm = Nothing
                End If
            Catch ex As Exception
                changeFrm = Nothing
            End Try

        End If

    End Sub

    Private Sub btnEnd2_Click(sender As Object, e As EventArgs) Handles btnEnd2.Click
        If Not IsNothing(timeStamps) Then

            If timeStamps.Count > 0 Then

                Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.letzter)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

                End If
            End If
        End If

    End Sub

    Private Sub btnEnd2_MouseHover(sender As Object, e As EventArgs) Handles btnEnd2.MouseHover
        Dim tmpDate As Date = getNextNavigationDate(ptNavigationButtons.letzter, True)
        ToolTipTS.Show(tmpDate.ToString, btnEnd, 2000)
    End Sub
End Class