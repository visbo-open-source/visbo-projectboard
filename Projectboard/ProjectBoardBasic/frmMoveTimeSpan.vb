Imports ProjectBoardDefinitions

Public Class frmMoveTimeSpan


    Private Sub moveToLeft_Click(sender As Object, e As EventArgs) Handles moveToLeft.Click

        Dim von As Integer, bis As Integer

        If showrangeleft > 1 Then
            von = showrangeleft - 1
            bis = showrangeright - 1
            Call awinChangeTimeSpan(von, bis)
        Else
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
        End If

    End Sub

    Private Sub moveToRight_Click(sender As Object, e As EventArgs) Handles moveToRight.Click
        Dim von As Integer, bis As Integer

        If showRangeRight < 2400 Then
            von = showRangeLeft + 1
            bis = showRangeRight + 1
            Call awinChangeTimeSpan(von, bis)
        Else
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
        End If
    End Sub

    Private Sub frmMoveTimeSpan_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If awinSettings.englishLanguage Then
            Me.Text = "Move Timespan"
        Else
            Me.Text = "Zeitraum verschieben"
        End If

    End Sub
End Class