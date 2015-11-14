Imports ProjectBoardDefinitions

Public Class frmEditWoerterbuch

    Private allStandardPhases As New Collection
    Private allUnknownPhases As New Collection

    Private allStandardMilestones As New Collection
    Private allUnknownMilestones As New Collection


    ''' <summary>
    ''' wird beim Laden des Formuars durchlaufen 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmEditWoerterbuch_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim phaseDef As clsPhasenDefinition
        Dim milestoneDef As clsMeilensteinDefinition

        
        ' Phasen ...
        ' jetzt werden die allStandard.. und allUnknown.. - Listen aufgebaut 
        For i = 1 To missingPhaseDefinitions.Count

            phaseDef = missingPhaseDefinitions.getPhaseDef(i)

            If Not IsNothing(phaseDef) Then
                If Not allUnknownPhases.Contains(phaseDef.name) Then
                    allUnknownPhases.Add(phaseDef.name, phaseDef.name)
                End If

            End If
        Next

        ' die Liste der Standard Bezeichnungen aufbauen 
        For i = 1 To PhaseDefinitions.Count
            phaseDef = PhaseDefinitions.getPhaseDef(i)

            If Not IsNothing(phaseDef) Then
                If Not allStandardPhases.Contains(phaseDef.name) Then
                    allStandardPhases.Add(phaseDef.name, phaseDef.name)
                End If
            End If
        Next


        ' Meilensteine  ...
        ' jetzt werden die allStandard.. und allUnknown.. - Listen aufgebaut 
        For i = 1 To missingMilestoneDefinitions.Count

            milestoneDef = missingMilestoneDefinitions.getMilestoneDef(i)

            If Not IsNothing(milestoneDef) Then
                If Not allUnknownMilestones.Contains(milestoneDef.name) Then
                    allUnknownMilestones.Add(milestoneDef.name, milestoneDef.name)
                End If

            End If
        Next

        ' die Liste der Standard Bezeichnungen aufbauen 
        For i = 1 To MilestoneDefinitions.Count
            milestoneDef = MilestoneDefinitions.getMilestoneDef(i)

            If Not IsNothing(milestoneDef) Then
                If Not allStandardMilestones.Contains(milestoneDef.name) Then
                    allStandardMilestones.Add(milestoneDef.name, milestoneDef.name)
                End If
            End If
        Next


        ' jetzt Radio-Button setzen 
        If missingPhaseDefinitions.Count = 0 And missingMilestoneDefinitions.Count > 0 Then
            Me.rdbListShowsMilestones.Checked = True
        Else
            Me.rdbListShowsPhases.Checked = True
        End If


    End Sub

    Private Sub rdbListShowsPhases_CheckedChanged(sender As Object, e As EventArgs) Handles rdbListShowsPhases.CheckedChanged


        If rdbListShowsPhases.Checked Then

            unknownList.Items.Clear()
            standardList.Items.Clear()

            ' die Liste der unbekannten Phasen Bezeichnungen aufbauen 
            For i = 1 To allUnknownPhases.Count
                unknownList.Items.Add(allUnknownPhases.Item(i))
            Next

            ' die Liste der Standard Bezeichnungen aufbauen 
            For i = 1 To allStandardPhases.Count
                standardList.Items.Add(allStandardPhases.Item(i))
            Next

        Else
            unknownList.Items.Clear()
            standardList.Items.Clear()

            ' die Liste der unbekannten Meilenstein Bezeichnungen aufbauen 
            For i = 1 To allUnknownMilestones.Count
                unknownList.Items.Add(allUnknownMilestones.Item(i))
            Next

            ' die Liste der Standard Bezeichnungen aufbauen 
            For i = 1 To allStandardMilestones.Count
                standardList.Items.Add(allStandardMilestones.Item(i))
            Next
        End If

        ' jetzt noch die Filter zurücksetzen 
        filterUnknown.Text = ""
        filterStandard.Text = ""

    End Sub

    Private Sub filterUnknown_TextChanged(sender As Object, e As EventArgs) Handles filterUnknown.TextChanged

        Dim suchstr As String = filterUnknown.Text
        Dim currentNames As New Collection

        If rdbListShowsPhases.Checked Then
            currentNames = allUnknownPhases
        Else
            currentNames = allUnknownMilestones
        End If


        If filterUnknown.Text = "" Then
            unknownList.Items.Clear()
            For Each s As String In currentNames
                unknownList.Items.Add(s)
            Next
        Else
            unknownList.Items.Clear()
            For Each s As String In currentNames
                If s.Contains(suchstr) Then
                    unknownList.Items.Add(s)
                End If
            Next
        End If

        editUnknownItem.Text = ""

    End Sub

    Private Sub filterStandard_TextChanged(sender As Object, e As EventArgs) Handles filterStandard.TextChanged

        Dim suchstr As String = filterStandard.Text
        Dim currentNames As New Collection

        If rdbListShowsPhases.Checked Then
            currentNames = allStandardPhases
        Else
            currentNames = allStandardMilestones
        End If


        If filterStandard.Text = "" Then
            standardList.Items.Clear()
            For Each s As String In currentNames
                standardList.Items.Add(s)
            Next
        Else
            standardList.Items.Clear()
            For Each s As String In currentNames
                If s.Contains(suchstr) Then
                    standardList.Items.Add(s)
                End If
            Next
        End If

        editUnknownItem.Text = ""

    End Sub

    Private Sub clearUnknownList_Click(sender As Object, e As EventArgs) Handles clearUnknownList.Click

        unknownList.SelectedItems.Clear()

    End Sub

    Private Sub clearStandardList_Click(sender As Object, e As EventArgs) Handles clearStandardList.Click

        standardList.SelectedItems.Clear()

    End Sub

    Private Sub unknownList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles unknownList.SelectedIndexChanged
        If unknownList.SelectedItems.Count = 1 And standardList.SelectedItems.Count = 0 Then
            editUnknownItem.Text = unknownList.SelectedItem.ToString
        Else
            editUnknownItem.Text = ""
        End If
    End Sub

    Private Sub standardList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles standardList.SelectedIndexChanged
        If standardList.SelectedItems.Count = 1 And unknownList.SelectedItems.Count = 0 Then
            editUnknownItem.Text = standardList.SelectedItem.ToString
        Else
            editUnknownItem.Text = ""
        End If
    End Sub
End Class