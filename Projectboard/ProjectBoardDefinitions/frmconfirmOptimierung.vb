Public Class frmconfirmOptimierung

    Private Sub ButtonJA_Click(sender As System.Object, e As System.EventArgs) Handles ButtonJA.Click

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()

    End Sub

    Private Sub ButtonNEIN_Click(sender As System.Object, e As System.EventArgs) Handles ButtonNEIN.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

End Class