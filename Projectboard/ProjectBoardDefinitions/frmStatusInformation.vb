Public Class frmStatusInformation

    Private Sub frmStatusInformation_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'Call awinDeleteProjectChildShapes(2)
        Call awinDeSelect()
        

    End Sub

    Private Sub frmStatusInformation_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmCoord(PTfrm.projInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.projInfo, PTpinfo.left) = Me.Left
    End Sub


    Private Sub frmStatusInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = CInt(frmCoord(PTfrm.projInfo, PTpinfo.top))
        Me.Left = CInt(frmCoord(PTfrm.projInfo, PTpinfo.left))

    End Sub
End Class