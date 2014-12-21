Public Class frmMppSettings

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With awinSettings
            shwProjectLine.Checked = .mppShowProjectLine
            notStrictly.Checked = .mppStrict
            shwAmpeln.Checked = .mppShowAmpel
            shwPhaseText.Checked = .mppShowPhName
            shwPhaseDate.Checked = .mppShowPhDate
            phaseFullyContained.Checked = .mppFullyContained
            ShwMilestoneText.Checked = .mppShowMsName
            ShwMilestoneDate.Checked = .mppShowMsDate
            shwVerticals.Checked = .mppVertikalesRaster
            shwLegend.Checked = .mppShowLegend

        End With
        

    End Sub

    

    Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click

        awinSettings.mppShowProjectLine = shwProjectLine.Checked
        awinSettings.mppStrict = Not notStrictly.Checked
        awinSettings.mppShowAmpel = shwAmpeln.Checked
        awinSettings.mppShowPhName = shwPhaseText.Checked
        awinSettings.mppShowPhDate = shwPhaseDate.Checked
        awinSettings.mppFullyContained = phaseFullyContained.Checked
        awinSettings.mppShowMsName = ShwMilestoneText.Checked
        awinSettings.mppShowMsDate = ShwMilestoneDate.Checked
        awinSettings.mppVertikalesRaster = shwVerticals.Checked
        awinSettings.mppShowLegend = shwLegend.Checked

        MyBase.Close()

    End Sub
End Class