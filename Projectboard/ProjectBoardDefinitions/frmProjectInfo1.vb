Imports ProjectBoardDefinitions
Public Class frmProjectInfo1

    Private Sub checkLanguageAndVisibility()
        If awinSettings.englishLanguage Then
            Me.lblCurrentVersion.Text = "current Version"
            Me.Text = "Profit/Loss Forecast"
        Else
            Me.lblCurrentVersion.Text = "aktuelle Version"
            Me.Text = "Gewinn/Verlust Prognose"
        End If
    End Sub


    Private Sub frmProjectInfo1_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        ' die globale Variable , die auf dieses Formular zeigt wird auf Nothing gesetzt  
        formProjectInfo1 = Nothing

    End Sub

    Private Sub frmProjectInfo1_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmCoord(PTfrm.projInfoPL, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.projInfoPL, PTpinfo.left) = Me.Left
    End Sub

    Private Sub frmProjectInfo1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Me.Top = CInt(frmCoord(PTfrm.projInfoPL, PTpinfo.top))
        'Me.Left = CInt(frmCoord(PTfrm.projInfoPL, PTpinfo.left))

        Call checkLanguageAndVisibility()

    End Sub
End Class