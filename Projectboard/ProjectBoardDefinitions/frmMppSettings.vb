Public Class frmMppSettings

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With awinSettings

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


        End With


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

        If awinSettings.mppSortiertDauer Then
            awinSettings.mppShowAllIfOne = True
        End If

        ' Änderung tk: geändert, weil sonst die Phasen nicht mehr ganz angezeigt werden ... 
        'awinSettings.mppFullyContained = awinSettings.mppSortiertDauer

        MyBase.Close()

    End Sub

    
End Class