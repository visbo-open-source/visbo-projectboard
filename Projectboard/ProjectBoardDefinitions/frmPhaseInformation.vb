Public Class frmPhaseInformation

    Private oldStart As Date, oldEnd As Date
    Private newStart As Date, newEnd As Date
    Private oldDauer As Integer, newDauer As Integer



    Private Sub frmPhaseInformation_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.phaseInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.phaseInfo, PTpinfo.left) = Me.Left

        Call awinDeleteMilestoneShapes(3)

    End Sub


    Private Sub frmPhaseInformation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = frmCoord(PTfrm.phaseInfo, PTpinfo.top)
        Me.Left = frmCoord(PTfrm.phaseInfo, PTpinfo.left)

        oldStart = CDate(phaseStart.Text)
        oldEnd = CDate(phaseEnde.Text)


    End Sub

    Private Sub phaseStart_GotFocus(sender As Object, e As EventArgs) Handles phaseStart.GotFocus
        appInstance.EnableEvents = False
        enableOnUpdate = False
    End Sub

    'Private Sub phaseStart_Leave(sender As Object, e As EventArgs) Handles phaseStart.Leave
    '    appInstance.EnableEvents = True
    '    enableOnUpdate = True
    '    Call MsgBox("Leave!")
    'End Sub

    Private Sub phaseStart_LostFocus(sender As Object, e As EventArgs) Handles phaseStart.LostFocus
        Dim validChange As Boolean = False
        Dim hproj As clsProjekt
        Dim cPhase As clsPhase


        Try
            hproj = ShowProjekte.getProject(projectName.Text)
            cPhase = hproj.getPhase(phaseName.Text)

            newStart = CDate(phaseStart.Text)
            newEnd = CDate(phaseEnde.Text)
            newDauer = calcDauerIndays(newStart, newEnd)
        Catch ex As Exception

        End Try

        ' jetzt wieder zurücksetzen der Event Behandlung 
        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class