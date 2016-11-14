Imports ProjectBoardDefinitions
Public Class frmTimeMachine

    Private nrSnapshots As Integer
    Private valueBeauftragung As Integer
    Private minmaxScales(1, 6) As Double
    Private necessary(6) As Boolean
    Private hproj As clsProjekt
    Private showAll As Boolean = False
    Private phaseList As Collection
    Private milestoneList As Collection
    Private typCollection As New Collection
    Private lastAmpel As Integer

    Private Sub frmTimeMachine_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' die beiden Buttons Home und ChangedPosition invisible setzen ..
        btnChangedPosition.Visible = False
        btnHome.Visible = False



    End Sub
End Class