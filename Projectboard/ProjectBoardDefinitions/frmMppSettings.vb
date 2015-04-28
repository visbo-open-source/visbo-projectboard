Public Class frmMppSettings
    ' ur: 20.04.2015: ???? hier gehts weiter
    Public calledfrom As String

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If calledfrom = "frmShowPlanElements" Then

            With awinSettings

                .eppExtendedMode = False

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                shwProjectLine.Checked = .mppShowProjectLine
                notStrictly.Checked = .mppShowAllIfOne
                shwAmpeln.Checked = .mppShowAmpel
                shwPhaseText.Checked = .mppShowPhName
                shwPhaseDate.Checked = .mppShowPhDate
                'phaseFullyContained.Checked = .mppFullyContained
                ShwMilestoneText.Checked = .mppShowMsName
                ShwMilestoneDate.Checked = .mppShowMsDate
                shwVerticals.Checked = .mppVertikalesRaster
                shwLegend.Checked = .mppShowLegend
                sortiertNachDauer.Checked = .mppSortiertDauer
                allOnOnePage.Checked = .mppOnePage
                shwExtendedMode.Checked = .mppExtendedMode

            End With
        ElseIf calledfrom = "frmSelectPPTTempl" Then

            With awinSettings

                .eppExtendedMode = True

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                shwProjectLine.Checked = .mppShowProjectLine
                notStrictly.Visible = False
                notStrictly.Checked = .mppShowAllIfOne
                shwAmpeln.Checked = .mppShowAmpel
                shwPhaseText.Checked = .mppShowPhName
                shwPhaseDate.Checked = .mppShowPhDate
                'phaseFullyContained.Checked = .mppFullyContained
                ShwMilestoneText.Checked = .mppShowMsName
                ShwMilestoneDate.Checked = .mppShowMsDate
                shwVerticals.Checked = .mppVertikalesRaster
                shwLegend.Checked = .mppShowLegend
                sortiertNachDauer.Visible = False
                'sortiertNachDauer.Checked = .mppSortiertDauer
                allOnOnePage.Visible = False
                'allOnOnePage.Checked = .mppOnePage
                shwExtendedMode.Visible = False
                'shwExtendedMode.Checked = .mppExtendedMode

                'ur: 21.04.2015: noch zu tun: zuvor alten Wert sichern
                shwExtendedMode.Checked = True
            End With
        End If

    End Sub



    Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click

        awinSettings.mppShowProjectLine = shwProjectLine.Checked
        awinSettings.mppShowAllIfOne = notStrictly.Checked
        awinSettings.mppShowAmpel = shwAmpeln.Checked
        awinSettings.mppShowPhName = shwPhaseText.Checked
        awinSettings.mppShowPhDate = shwPhaseDate.Checked
        awinSettings.mppShowMsName = ShwMilestoneText.Checked
        awinSettings.mppShowMsDate = ShwMilestoneDate.Checked
        awinSettings.mppVertikalesRaster = shwVerticals.Checked
        awinSettings.mppShowLegend = shwLegend.Checked
        awinSettings.mppOnePage = allOnOnePage.Checked
        awinSettings.mppSortiertDauer = sortiertNachDauer.Checked
        awinSettings.mppExtendedMode = shwExtendedMode.Checked

        If awinSettings.mppSortiertDauer Then
            awinSettings.mppShowAllIfOne = True
        End If

        ' Änderung tk: geändert, weil sonst die Phasen nicht mehr ganz angezeigt werden ... 
        'awinSettings.mppFullyContained = awinSettings.mppSortiertDauer

        MyBase.Close()

    End Sub


    Private Sub notStrictly_CheckedChanged(sender As Object, e As EventArgs) Handles notStrictly.CheckedChanged

    End Sub

    Private Sub shwExtendedMode_CheckedChanged(sender As Object, e As EventArgs) Handles shwExtendedMode.CheckedChanged

    End Sub
End Class