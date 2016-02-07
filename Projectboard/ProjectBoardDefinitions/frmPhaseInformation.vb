Public Class frmPhaseInformation


    Public phaseNameID As String
    Public curProject As clsProjekt

    Private Sub frmPhaseInformation_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.phaseInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.phaseInfo, PTpinfo.left) = Me.Left

        'Call awinDeleteProjectChildShapes(3)
        Call awinDeSelect()

    End Sub


    Private Sub frmPhaseInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = CInt(frmCoord(PTfrm.phaseInfo, PTpinfo.top))
        Me.Left = CInt(frmCoord(PTfrm.phaseInfo, PTpinfo.left))

        Me.showOrigItem.Checked = awinSettings.showOrigName

    End Sub

    

    'Private Sub phaseStart_Leave(sender As Object, e As EventArgs) Handles phaseStart.Leave
    '    appInstance.EnableEvents = True
    '    enableOnUpdate = True
    '    Call MsgBox("Leave!")
    'End Sub

    

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    ''' <summary>
    ''' zeigt den urspünglichen Phasen-Namen aus Rplan oder anderem PM-System an 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub showOrigItem_CheckedChanged(sender As Object, e As EventArgs) Handles showOrigItem.CheckedChanged
        Dim tmpNode As clsHierarchyNode

        awinSettings.showOrigName = showOrigItem.Checked

        If showOrigItem.Checked = True Then
            tmpNode = curProject.hierarchy.nodeItem(phaseNameID)
            If Not IsNothing(tmpNode) Then
                phaseName.Text = tmpNode.origName
            Else
                phaseName.Text = elemNameOfElemID(phaseNameID)
            End If
        Else
            phaseName.Text = elemNameOfElemID(phaseNameID)
        End If
    End Sub
End Class