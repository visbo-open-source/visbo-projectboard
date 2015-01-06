Public Class frmStatusInformation

    Private Sub frmStatusInformation_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        frmCoord(PTfrm.projInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.projInfo, PTpinfo.left) = Me.Left

        Call awinDeleteProjectChildShapes(2)
        

    End Sub


    Private Sub frmStatusInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = CInt(frmCoord(PTfrm.projInfo, PTpinfo.top))
        Me.Left = CInt(frmCoord(PTfrm.projInfo, PTpinfo.left))

    End Sub
End Class