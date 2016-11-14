Public Class frmPPTTimeMachine
    Private currentIndex As Integer = -1
    Private currentTimestamp As Date = Date.MinValue
    Private timeStampsArray() As Date = Nothing
    ''' <summary>
    ''' Laden der Time-Machine
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmPPTTimeMachine_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If noDBAccessInPPT Then
            Call MsgBox("kein Datenbank Zugriff ... Abbruch ...")
            MyBase.Close()
        Else
            ' die beiden Buttons Home und ChangedPosition invisible setzen ..
            btnChangedPosition.Visible = False
            btnHome.Visible = False

            ' die Navigation de-aktivieren 

            ' gibt es überhaupt TimeStamps ? 
            timeStampsArray = smartSlideLists.getArrayOfTS()
            If Not IsNothing(timeStampsArray) Then
                If timeStampsArray.Length >= 1 Then

                    ' die Navigation enablen  
                    btnBack.Enabled = False
                    btnFastBack.Enabled = False
                    btnStart.Enabled = True
                    btnForward.Enabled = False
                    btnFastForward.Enabled = False
                    btnEnd.Enabled = True
                    txtboxCurrentDate.Enabled = True
                    txtboxCurrentDate.Text = ""
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
                End If
            End If

            ' jetzt wird der aktuelle Stand geholt .. und mit den aktuellen Termindaten verglichen ...

            ' wenn der Stand unterschiedlich ist, werden die Moving Forward-/Backward Buttons entsprechend visible gesetzt 

        End If

    End Sub

    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        currentIndex = timeStampsArray.Length - 1
        currentTimestamp = timeStampsArray(currentIndex)

        ' alle zur Home-Position schicken ...
        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            If isRelevantShape(tmpShape) Then
                Call sendToTimeStampPosition(tmpShape, currentTimestamp)
            End If
        Next

    End Sub
End Class