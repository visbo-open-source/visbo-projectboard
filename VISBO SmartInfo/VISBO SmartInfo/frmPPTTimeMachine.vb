Public Class frmPPTTimeMachine

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

            ' jetzt wird der aktuelle Stand geholt .. und mit den aktuellen Termindaten verglichen ...

            ' wenn der Stand unterschiedlich ist, werden die Moving Forward-/Backward Buttons entsprechend visible gesetzt 

        End If

    End Sub
End Class