Public Class frmMppSettings

    Public calledfrom As String

    Private Sub frmMppSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With awinSettings
            shwProjectLine.Checked = .mppShowProjectLine
            notStrictly.Checked = .mppShowAllIfOne
            shwAmpeln.Checked = .mppShowAmpel
            useOriginalNames.Checked = .mppUseOriginalNames

            shwPhaseText.Checked = .mppShowPhName
            shwPhaseDate.Checked = .mppShowPhDate
            useAbbrev.Checked = .mppUseAbbreviation
            ShwMilestoneText.Checked = .mppShowMsName
            ShwMilestoneDate.Checked = .mppShowMsDate
            KwInMilestone.Checked = .mppKwInMilestone

            shwVerticals.Checked = .mppVertikalesRaster
            shwLegend.Checked = .mppShowLegend
            sortiertNachDauer.Checked = .mppSortiertDauer
            shwHorizontals.Checked = .mppShowHorizontals
            allOnOnePage.Checked = .mppOnePage
            shwExtendedMode.Checked = .mppExtendedMode

            If .mppSortiertDauer Then
                .mppShowAllIfOne = True
            End If

            '.eppExtendedMode = False
        End With


        If calledfrom = "frmShowPlanElements" Then

            ' alle Elemente anzeigen 

        ElseIf calledfrom = "frmSelectPPTTempl" Then

            With awinSettings

                '.eppExtendedMode = True

                ' was soll nicht visible sein 
                notStrictly.Visible = False
                sortiertNachDauer.Visible = False
                allOnOnePage.Visible = False
                shwExtendedMode.Visible = False
                'ur: 21.04.2015: noch zu tun: zuvor alten Wert sichern
                shwExtendedMode.Checked = True

                ' den UseOriginal NAmes Button hochschieben 
                useOriginalNames.Top = notStrictly.Top
            End With

        ElseIf calledfrom = "frmBHTC" Then

            With awinSettings

                ' die folgenden werden im MS Project AddIn gar nicht gezeigt 
                ' sie müssen auch entsprechend auf die für den Project AddIn sinnvllen werte gesetzt werden 
                '
                shwProjectLine.Visible = False
                shwAmpeln.Visible = False
                notStrictly.Visible = False
                useOriginalNames.Visible = False
                sortiertNachDauer.Visible = False
                shwExtendedMode.Visible = False
                

                ' jetzt müssen die Checkboxes und der OK-Button noch hochgeschoben werden 
                ' ausserdem die Höhe des Formulars verändert werden 
                Dim offset As Integer = shwPhaseText.Top - shwProjectLine.Top
                shwPhaseText.Top = shwPhaseText.Top - offset
                shwPhaseDate.Top = shwPhaseDate.Top - offset

                ShwMilestoneText.Top = ShwMilestoneText.Top - offset
                ShwMilestoneDate.Top = ShwMilestoneDate.Top - offset
                KwInMilestone.Top = KwInMilestone.Top - offset

                useAbbrev.Top = useAbbrev.Top - offset

                shwVerticals.Top = shwVerticals.Top - offset
                shwHorizontals.Top = shwHorizontals.Top - offset

                shwLegend.Top = shwLegend.Top - offset
                allOnOnePage.Top = allOnOnePage.Top - offset

                okButton.Top = okButton.Top - offset

                Me.Height = Me.Height - offset


            End With


        End If

    End Sub



    Private Sub okButton_Click(sender As Object, e As EventArgs) Handles okButton.Click


        awinSettings.mppShowProjectLine = shwProjectLine.Checked
        awinSettings.mppShowAllIfOne = notStrictly.Checked
        awinSettings.mppShowAmpel = shwAmpeln.Checked
        awinSettings.mppUseAbbreviation = useOriginalNames.Checked

        awinSettings.mppShowPhName = shwPhaseText.Checked
        awinSettings.mppShowPhDate = shwPhaseDate.Checked
        awinSettings.mppShowMsName = ShwMilestoneText.Checked
        awinSettings.mppShowMsDate = ShwMilestoneDate.Checked
        awinSettings.mppUseAbbreviation = useAbbrev.Checked
        awinSettings.mppKwInMilestone = KwInMilestone.Checked


        awinSettings.mppVertikalesRaster = shwVerticals.Checked
        awinSettings.mppShowHorizontals = shwHorizontals.Checked
        awinSettings.mppShowLegend = shwLegend.Checked
        awinSettings.mppOnePage = allOnOnePage.Checked

        awinSettings.mppSortiertDauer = sortiertNachDauer.Checked
        awinSettings.mppExtendedMode = shwExtendedMode.Checked

        If awinSettings.mppSortiertDauer Then
            awinSettings.mppShowAllIfOne = True
        End If


        MyBase.Close()

    End Sub


    Private Sub notStrictly_CheckedChanged(sender As Object, e As EventArgs) Handles notStrictly.CheckedChanged

    End Sub

    Private Sub shwExtendedMode_CheckedChanged(sender As Object, e As EventArgs) Handles shwExtendedMode.CheckedChanged

    End Sub
End Class