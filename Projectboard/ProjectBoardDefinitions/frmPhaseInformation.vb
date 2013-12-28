Public Class frmPhaseInformation

    Private Sub frmPhaseInformation_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.phaseInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.phaseInfo, PTpinfo.left) = Me.Left

        Call awinDeleteMilestoneShapes(3)

    End Sub


    Private Sub frmPhaseInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = frmCoord(PTfrm.phaseInfo, PTpinfo.top)
        Me.Left = frmCoord(PTfrm.phaseInfo, PTpinfo.left)


    End Sub

End Class