
Imports ProjectBoardDefinitions
Public Class frmBetterWorseSettings

    Private chartType As Integer
    Private auswahl As Integer

    Private Sub frmBetterWorseSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub timeTolerance_TextChanged(sender As Object, e As EventArgs) Handles timeTolerance.TextChanged

        If Me.RBvglRel.Checked Then
            awinSettings.timeToleranzRel = CDbl(timeTolerance.Text) / 100
        Else
            awinSettings.timeToleranzAbs = CDbl(timeTolerance.Text) / 100
        End If

    End Sub

    Private Sub costTolerance_TextChanged(sender As Object, e As EventArgs) Handles costTolerance.TextChanged

        If Me.RBvglRel.Checked Then
            awinSettings.costToleranzRel = CDbl(costTolerance.Text) / 100
        Else
            awinSettings.costToleranzAbs = CDbl(costTolerance.Text) / 100
        End If


    End Sub

    Private Sub RBvglRel_CheckedChanged(sender As Object, e As EventArgs) Handles RBvglRel.CheckedChanged

        Dim tmpVal As Double
        tmpVal = awinSettings.costToleranzRel * 100
        costTolerance.Text = tmpVal.ToString("##0.#")
        CostToleranz.Text = "Kosten Toleranz in %"


        tmpVal = awinSettings.timeToleranzRel * 100
        timeTolerance.Text = tmpVal.ToString("##0.#")
        timeToleranz.Text = "Zeit Toleranz in %"
    End Sub

    Private Sub RBvglAbs_CheckedChanged(sender As Object, e As EventArgs) Handles RBvglAbs.CheckedChanged

        Dim tmpVal As Double
        tmpVal = awinSettings.costToleranzAbs
        costTolerance.Text = tmpVal.ToString("##0")
        CostToleranz.Text = "Kosten Toleranz in T€"

        tmpVal = awinSettings.timeToleranzAbs
        timeTolerance.Text = tmpVal.ToString("##0")
        timeToleranz.Text = "Zeit Toleranz in Tagen"

    End Sub

    Private Sub RBvglB_CheckedChanged(sender As Object, e As EventArgs) Handles RBvglB.CheckedChanged

    End Sub

    Private Sub CBendOfP_CheckedChanged(sender As Object, e As EventArgs) Handles CBendOfP.CheckedChanged

        auswahl = 0
    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

    End Sub
End Class