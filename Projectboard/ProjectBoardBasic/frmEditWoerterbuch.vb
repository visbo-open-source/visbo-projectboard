Imports ProjectBoardDefinitions

Public Class frmEditWoerterbuch

    'Private allStandardPhases As New Collection
    'Private allUnknownPhases As New Collection

    'Private allStandardMilestones As New Collection
    'Private allUnknownMilestones As New Collection
    ' wird benutzt , um beim Listen Aktualisieren eine Aktion zu triggern oder eben nicht 
    Private eventsShouldFire As Boolean = True


    ''' <summary>
    ''' wird beim Laden des Formuars durchlaufen 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmEditWoerterbuch_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        ' jetzt Radio-Button setzen 
        If missingPhaseDefinitions.Count = 0 And missingMilestoneDefinitions.Count > 0 Then
            Me.rdbListShowsMilestones.Checked = True
        Else
            Me.rdbListShowsPhases.Checked = True
        End If


    End Sub

    Private Sub rdbListShowsPhases_CheckedChanged(sender As Object, e As EventArgs) Handles rdbListShowsPhases.CheckedChanged

        Dim tmpPhDef As clsPhasenDefinition
        Dim tmpMsDef As clsMeilensteinDefinition

        If rdbListShowsPhases.Checked Then

            unknownList.Items.Clear()
            standardList.Items.Clear()

            ' die Liste der unbekannten Phasen Bezeichnungen aufbauen 
            For i = 1 To missingPhaseDefinitions.Count
                tmpPhDef = missingPhaseDefinitions.getPhaseDef(i)
                If Not IsNothing(tmpPhDef) Then
                    unknownList.Items.Add(tmpPhDef.name)
                End If

            Next

            ' die Liste der Standard Bezeichnungen aufbauen 
            For i = 1 To PhaseDefinitions.Count
                tmpPhDef = PhaseDefinitions.getPhaseDef(i)
                If Not IsNothing(tmpPhDef) Then
                    standardList.Items.Add(tmpPhDef.name)
                End If

            Next

        Else
            unknownList.Items.Clear()
            standardList.Items.Clear()

            ' die Liste der unbekannten Meilenstein Bezeichnungen aufbauen 
            For i = 1 To missingMilestoneDefinitions.Count
                tmpMsDef = missingMilestoneDefinitions.getMilestoneDef(i)
                If Not IsNothing(tmpMsDef) Then
                    unknownList.Items.Add(tmpMsDef.name)
                End If
            Next

            ' die Liste der Standard Bezeichnungen aufbauen 
            For i = 1 To MilestoneDefinitions.Count
                tmpMsDef = MilestoneDefinitions.getMilestoneDef(i)
                If Not IsNothing(tmpMsDef) Then
                    standardList.Items.Add(tmpMsDef.name)
                End If
            Next
        End If

        ' jetzt noch die Filter zurücksetzen 
        filterUnknown.Text = ""
        filterStandard.Text = ""

    End Sub

    Private Sub filterUnknown_TextChanged(sender As Object, e As EventArgs) Handles filterUnknown.TextChanged

        Dim suchstr As String = filterUnknown.Text
        Dim tmpName As String

        unknownList.Items.Clear()

        If rdbListShowsPhases.Checked Then

            For i = 1 To missingPhaseDefinitions.Count
                tmpName = missingPhaseDefinitions.getPhaseDef(i).name
                If filterUnknown.Text = "" Then
                    unknownList.Items.Add(tmpName)
                Else
                    If tmpName.Contains(suchstr) Then
                        unknownList.Items.Add(tmpName)
                    End If
                End If
            Next

        Else
            ' Meilensteine
            For i = 1 To missingMilestoneDefinitions.Count
                tmpName = missingMilestoneDefinitions.getMilestoneDef(i).name
                If filterUnknown.Text = "" Then
                    unknownList.Items.Add(tmpName)
                Else
                    If tmpName.Contains(suchstr) Then
                        unknownList.Items.Add(tmpName)
                    End If
                End If
            Next

        End If

        editUnknownItem.Text = ""

    End Sub

    Private Sub filterStandard_TextChanged(sender As Object, e As EventArgs) Handles filterStandard.TextChanged

        Dim suchstr As String = filterStandard.Text
        Dim tmpName As String

        standardList.Items.Clear()

        If rdbListShowsPhases.Checked Then

            For i = 1 To PhaseDefinitions.Count
                tmpName = PhaseDefinitions.getPhaseDef(i).name
                If filterStandard.Text = "" Then
                    standardList.Items.Add(tmpName)
                Else
                    If tmpName.Contains(suchstr) Then
                        standardList.Items.Add(tmpName)
                    End If
                End If
            Next

        Else
            ' Meilensteine
            For i = 1 To MilestoneDefinitions.Count
                tmpName = MilestoneDefinitions.getMilestoneDef(i).name
                If filterStandard.Text = "" Then
                    standardList.Items.Add(tmpName)
                Else
                    If tmpName.Contains(suchstr) Then
                        standardList.Items.Add(tmpName)
                    End If
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

        If eventsShouldFire Then
            If unknownList.SelectedItems.Count = 1 And standardList.SelectedItems.Count = 0 Then
                editUnknownItem.Text = unknownList.SelectedItem.ToString
            Else
                editUnknownItem.Text = ""
            End If
        End If

        
    End Sub

    Private Sub standardList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles standardList.SelectedIndexChanged

        If eventsShouldFire Then
            If standardList.SelectedItems.Count = 1 And unknownList.SelectedItems.Count = 0 Then
                editUnknownItem.Text = standardList.SelectedItem.ToString
            Else
                editUnknownItem.Text = ""
            End If
        
        End If
        
    End Sub


    ''' <summary>
    ''' visualisiert die selektierten Elemente aus der Unknown-List bzw. aus der Standard-List
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub visualizeElements()

        Dim selectedphases As New Collection
        Dim selectedMilestones As New Collection
        Dim tmpName As String
        Dim deleteOtherShapes As Boolean = True
        Dim numberIt As Boolean = False
        Dim alleFarben As Integer = 4


        If rdbListShowsPhases.Checked Then
            ' es werden Phasen visualisiert 

            ' Aufbau der Phasenliste aus den selektierten Elementen der linken Liste  
            For i = 1 To unknownList.SelectedItems.Count
                tmpName = CStr(unknownList.SelectedItems.Item(i - 1)).Trim
                If Not selectedphases.Contains(tmpName) Then
                    selectedphases.Add(tmpName, tmpName)
                End If
            Next

            ' Erweiterung  der Phasenliste um die selektierten Elemente der rechten  Liste (Standard-Liste)   
            For i = 1 To standardList.SelectedItems.Count
                tmpName = CStr(standardList.SelectedItems.Item(i - 1)).Trim
                If Not selectedphases.Contains(tmpName) Then
                    selectedphases.Add(tmpName, tmpName)
                End If
            Next

            ' Phasen sollen nicht nummeriert werden 
            ' zuvor gezeichnete Phasen sollen gelöscht werden - oder auch nicht
            If selectedphases.Count > 0 Then
                Call awinZeichnePhasen(selectedphases, numberIt, deleteOtherShapes)
            Else
                Call MsgBox("bitte mindestens ein Element aus der linken und/oder rechten Liste auswählen")
            End If


        Else
            ' es werden Meilensteine visualisiert 

            ' Aufbau der Meilensteinliste aus den selektierten Elementen der linken Liste  
            For i = 1 To unknownList.SelectedItems.Count
                tmpName = CStr(unknownList.SelectedItems.Item(i - 1)).Trim
                If Not selectedMilestones.Contains(tmpName) Then
                    selectedMilestones.Add(tmpName, tmpName)
                End If
            Next

            ' Erweiterung  der Phasenliste um die selektierten Elemente der rechten  Liste (Standard-Liste)   
            For i = 1 To standardList.SelectedItems.Count
                tmpName = CStr(standardList.SelectedItems.Item(i - 1)).Trim
                If Not selectedMilestones.Contains(tmpName) Then
                    selectedMilestones.Add(tmpName, tmpName)
                End If
            Next

            ' Meilensteine sollen nicht nummeriert werden 
            ' zuvor gezeichnete Meilensteine sollen gelöscht werden - oder auch nicht
            If selectedMilestones.Count > 0 Then
                Call awinZeichneMilestones(selectedMilestones, alleFarben, numberIt, deleteOtherShapes)
            Else
                Call MsgBox("bitte mindestens ein Element aus der linken und/oder rechten Liste auswählen")
            End If

        End If

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub makeItemToBeStandard()

        Dim itemText As String
        Dim editText As String

        eventsShouldFire = False

        ' Vorbedingung prüfen: es darf kein Element in der Liste der Standard-Bezeichnungen selektiert sein ! 
        If standardList.SelectedItems.Count > 0 Then
            Call MsgBox("für diese Aktion darf kein Standard-Element selektiert sein !")
        Else
            ' die grundsätzliche Vorbedingung ist erfüllt: keine Selektion in der Standard-Liste 
            If unknownList.SelectedItems.Count = 0 Then

                Call MsgBox("bitte mindestens einen Eintrag aus der Liste der unbekannten Bezeichnungen selektieren")

            ElseIf unknownList.SelectedItems.Count = 1 Then
                ' es ist nur ein Element selektiert, d.h es könnte auch sein, dass es eine Änderung des Eintrages im EditUnknownItem gab

                editText = editUnknownItem.Text.Trim
                itemText = CStr(unknownList.SelectedItem).Trim

                ' falls der Anwender den Text zur leeren Zeichenkette gemacht hat ... 
                If editText.Length = 0 Then
                    Call MsgBox("leere Zeichenkette nicht zulässig ...")
                Else
                    ' jetzt ist so weit alles i.O - es muss geprüft werden, ob der itemtext geändert wurde oder gleich geblieben ist
                    ' wenn er gleich geblieben ist, dann wird das missingElement zum Standard gemacht und aus der Lsite der unknow gelöscht
                    ' wenn es sich geändert hat, dann wird das veränderte missingElement zum standard gemacht, aber das unveränderte bleibt in der Liste der Unknown


                    If rdbListShowsPhases.Checked Then
                        ' eine neue Phasen-Definition aufnehmen 

                        If editText = itemText Then
                            Try
                                ' die PhaseDef in die Standard-Liste aufnehmen 
                                ' aus der Liste der unbekannten Phasenbezeichnungen löschen 

                                Dim tmpPhDef As clsPhasenDefinition = missingPhaseDefinitions.getPhaseDef(itemText)
                                PhaseDefinitions.Add(tmpPhDef)
                                standardList.Items.Add(itemText)

                                missingPhaseDefinitions.remove(itemText)
                                unknownList.Items.Remove(itemText)

                            Catch ex As Exception

                            End Try

                        Else
                            Try
                                ' die geänderte PhaseDef aufnehmen, aber die alte in den missing noch nicht löschen ! 
                                ' dazu muss erst noch eine Abbildungsregel aufgebaut werden 

                                Dim tmpPhDef As New clsPhasenDefinition
                                tmpPhDef.copyFrom(missingPhaseDefinitions.getPhaseDef(itemText), editText)
                                PhaseDefinitions.Add(tmpPhDef)
                                standardList.Items.Add(editText)

                            Catch ex As Exception

                            End Try
                        End If
                    Else
                        ' eine neue Meilenstein Definition aufnehmen 

                        If editText = itemText Then
                            Try
                                ' die MilestoneDef in die Standard-Liste aufnehmen 
                                ' aus der Liste der unbekannten Milestonebezeichnungen löschen 

                                Dim tmpMsDef As clsMeilensteinDefinition = missingMilestoneDefinitions.getMilestoneDef(itemText)
                                MilestoneDefinitions.Add(tmpMsDef)
                                missingMilestoneDefinitions.remove(itemText)
                                unknownList.Items.Remove(itemText)


                            Catch ex As Exception

                            End Try

                        Else
                            Try
                                ' die geänderte MilestoneDef aufnehmen, aber die alte in den missing noch nicht löschen ! 
                                ' dazu muss erst noch eine Abbildungsregel aufgebaut werden 

                                Dim tmpMsDef As New clsMeilensteinDefinition
                                tmpMsDef.copyFrom(missingMilestoneDefinitions.getMilestoneDef(itemText), editText)
                                MilestoneDefinitions.Add(tmpMsDef)
                                standardList.Items.Add(editText)

                            Catch ex As Exception

                            End Try
                        End If
                    End If
                End If



            Else
                Dim todoList As New Collection

                For i = 1 To unknownList.SelectedItems.Count
                    itemText = CStr(unknownList.SelectedItems.Item(i - 1)).Trim
                    If rdbListShowsPhases.Checked Then
                        ' eine neue Phasen-Definition aufnehmen 
                        ' den Eintrag aus missingPhaseDefinitions rausnehmen 

                        Try
                            Dim tmpPhDef As clsPhasenDefinition = missingPhaseDefinitions.getPhaseDef(itemText)
                            PhaseDefinitions.Add(tmpPhDef)
                            standardList.Items.Add(itemText)

                            missingPhaseDefinitions.remove(itemText)
                            todoList.Add(itemText)

                        Catch ex As Exception

                        End Try


                    Else
                        ' eine neue Meilenstein Definition aufnehmen 
                        ' den Eintrag aus missingMilestoneDefinitions rausnehmen 

                        Try

                            Dim tmpMsDef As clsMeilensteinDefinition = missingMilestoneDefinitions.getMilestoneDef(itemText)
                            MilestoneDefinitions.Add(tmpMsDef)
                            standardList.Items.Add(itemText)

                            missingMilestoneDefinitions.remove(itemText)
                            todoList.Add(itemText)

                        Catch ex As Exception

                        End Try

                    End If
                Next

                ' jetzt löschen 
                For i = 1 To todoList.Count
                    itemText = CStr(todoList.Item(i)).Trim
                    unknownList.Items.Remove(itemText)
                Next
            End If


        End If

        eventsShouldFire = True

    End Sub

    

    Private Sub setItemToBeKnown_Click(sender As Object, e As EventArgs) Handles setItemToBeKnown.Click

        Call makeItemToBeStandard()

    End Sub

    ''' <summary>
    ''' nimmt das Item aus derListe der bekannten Standard-Namen heraus; muss aber sicherstellen, 
    ''' dass alle Wörterbuch Abbildungen auf dieses Element auch rausgenommen werden 
    ''' Vorbedingung: bei den unknownlist Elementen darf nichts selektiert sein 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub setItemToBeUnknown_Click(sender As Object, e As EventArgs) Handles setItemToBeUnknown.Click

        Dim itemText As String
        Dim editText As String

        eventsShouldFire = False

        If unknownList.SelectedItems.Count > 0 Then

            Call MsgBox("es darf in der Liste der unbekannten Bezeichnungen kein Eintrag selektiert sein ...")

        ElseIf standardList.SelectedItems.Count = 1 Then

            itemText = CStr(standardList.SelectedItem).Trim
            editText = editUnknownItem.Text.Trim

            If itemText = editText Then

                If rdbListShowsPhases.Checked Then

                    ' als erstes prüfen, ob es für dieses Element Einträge im Wörterbuch gibt ..
                    Dim anzEintraege As Integer = phaseMappings.getAnzahlMappingsFor(itemText)

                    Dim ok As Boolean = True
                    If anzEintraege > 0 Then
                        Call MsgBox(itemText & " hat " & anzEintraege & " im Wörterbuch. Trotzdem fortfahren?", MsgBoxStyle.OkCancel)
                        If MsgBoxResult.Ok Then
                            phaseMappings.removeStdName(itemText)
                            ok = True
                        Else
                            ok = False
                        End If
                    End If

                    If ok Then
                        Dim tmpPhDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(itemText)
                        missingPhaseDefinitions.Add(tmpPhDef)
                        unknownList.Items.Add(itemText)

                        PhaseDefinitions.remove(itemText)
                        standardList.Items.Remove(itemText)
                    End If


                Else

                    ' als erstes prüfen, ob es für dieses Element Einträge im Wörterbuch gibt ..
                    Dim anzEintraege As Integer = milestoneMappings.getAnzahlMappingsFor(itemText)

                    Dim ok As Boolean = True
                    If anzEintraege > 0 Then
                        Call MsgBox(itemText & " hat " & anzEintraege & " im Wörterbuch. Trotzdem fortfahren?", MsgBoxStyle.OkCancel)
                        If MsgBoxResult.Ok Then
                            milestoneMappings.removeStdName(itemText)
                            ok = True
                        Else
                            ok = False
                        End If
                    End If

                    If ok Then
                        Dim tmpMsDef As clsMeilensteinDefinition = MilestoneDefinitions.getMilestoneDef(itemText)

                        If Not missingMilestoneDefinitions.Contains(tmpMsDef.name) Then
                            missingMilestoneDefinitions.Add(tmpMsDef)
                        End If

                        unknownList.Items.Add(itemText)

                        MilestoneDefinitions.remove(itemText)
                        standardList.Items.Remove(itemText)
                    End If


                End If

            Else
                Call MsgBox("unklar, was zu tun ist. Bitte nicht den Text ändern, bevor ein Element aus der Liste der Standard-Bezeichnungen entfernt wird")
            End If

        Else

            Dim todoList As New Collection
            For i = 1 To standardList.SelectedItems.Count

                itemText = CStr(standardList.SelectedItems.Item(i - 1)).Trim
                If rdbListShowsPhases.Checked Then

                    Dim anzEintraege As Integer = phaseMappings.getAnzahlMappingsFor(itemText)

                    Dim ok As Boolean = True
                    If anzEintraege > 0 Then
                        Call MsgBox(itemText & " hat " & anzEintraege & " im Wörterbuch. Trotzdem fortfahren?", MsgBoxStyle.OkCancel)
                        If MsgBoxResult.Ok Then
                            phaseMappings.removeStdName(itemText)
                            ok = True
                        Else
                            ok = False
                        End If
                    End If

                    If ok Then
                        Dim tmpPhDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(itemText)

                        missingPhaseDefinitions.Add(tmpPhDef)
                        unknownList.Items.Add(itemText)

                        PhaseDefinitions.remove(itemText)
                        todoList.Add(itemText)
                    End If


                Else

                    ' als erstes prüfen, ob es für dieses Element Einträge im Wörterbuch gibt ..
                    Dim anzEintraege As Integer = milestoneMappings.getAnzahlMappingsFor(itemText)

                    Dim ok As Boolean = True
                    If anzEintraege > 0 Then
                        Call MsgBox(itemText & " hat " & anzEintraege & " im Wörterbuch. Trotzdem fortfahren?", MsgBoxStyle.OkCancel)
                        If MsgBoxResult.Ok Then
                            milestoneMappings.removeStdName(itemText)
                            ok = True
                        Else
                            ok = False
                        End If
                    End If

                    If ok Then
                        Dim tmpMsDef As clsMeilensteinDefinition = MilestoneDefinitions.getMilestoneDef(itemText)
                        missingMilestoneDefinitions.Add(tmpMsDef)
                        unknownList.Items.Add(itemText)


                        MilestoneDefinitions.remove(itemText)
                        todoList.Add(itemText)

                    End If


                End If
            Next

            For i = 1 To todoList.Count
                itemText = CStr(todoList.Item(i))
                standardList.Items.Remove(itemText)
            Next

        End If

        eventsShouldFire = True
        

    End Sub

    ''' <summary>
    ''' fügt den oder die selektierten Namen in die Liste der zu ignorierenden Elemente ein
    ''' Vorbedingung: kein Element aus der Standard-Liste darfselektiert sein 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ignoreButton_Click(sender As Object, e As EventArgs) Handles ignoreButton.Click


        eventsShouldFire = False

        If standardList.SelectedItems.Count > 0 Then
            Call MsgBox("in der Liste der bekannten Namen darf kein Eintrag selektiert sein ...")
        Else
            If unknownList.SelectedItems.Count > 0 Then

                If rdbListShowsPhases.Checked Then

                    For Each Obj As Object In unknownList.SelectedItems
                        Try
                            Dim itemName As String = CStr(Obj)
                            phaseMappings.addIgnoreName(itemName)
                            missingPhaseDefinitions.remove(itemName)
                            unknownList.Items.Remove(itemName)

                        Catch ex As Exception

                        End Try
                    Next

                Else

                    For Each Obj As Object In unknownList.SelectedItems
                        Try
                            Dim itemName As String = CStr(Obj)
                            milestoneMappings.addIgnoreName(itemName)
                            missingMilestoneDefinitions.remove(itemName)
                            unknownList.Items.Remove(itemName)
                        Catch ex As Exception

                        End Try
                    Next

                End If

            Else
                Call MsgBox("wenigstens ein Element aus der Liste der unbekannten Bezeichnungen selektieren ...")
            End If
        End If

        eventsShouldFire = True

    End Sub

    ''' <summary>
    ''' fügt eine oder mehrere Ersetzungs-Regeln zum wörterbuch hinzu 
    ''' Vorbedingung: es müssen in jeder der Listen mindestens ein Eintrag selektiert sein 
    ''' der EditText darf nicht anders sein 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub addRulesToDictionary_Click(sender As Object, e As EventArgs) Handles addRulesToDictionary.Click

        Dim itemName As String
        Dim stdName As String
        Dim todoList As New Collection


        eventsShouldFire = False

        If unknownList.SelectedItems.Count > 0 And standardList.SelectedItems.Count > 0 Then

            stdName = CStr(standardList.SelectedItem).Trim
            For Each obj As Object In unknownList.SelectedItems
                itemName = CStr(obj)
                If rdbListShowsPhases.Checked Then

                    Try
                        phaseMappings.addSynonym(itemName, stdName)
                        missingPhaseDefinitions.remove(itemName)
                        todoList.Add(itemName)

                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try


                Else
                    Try
                        milestoneMappings.addSynonym(itemName, stdName)
                        missingMilestoneDefinitions.remove(itemName)
                        todoList.Add(itemName)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                End If
            Next

            For i = 1 To todoList.Count
                itemName = CStr(todoList.Item(i))
                unknownList.Items.Remove(itemName)
            Next

        Else
            Call MsgBox(" es muss in beiden Listen ein Eintrag selektiert sein ...")
        End If

        ' die Standard-Name jetzt de-selektieren 
        standardList.SelectedItems.Clear()

        eventsShouldFire = True
        
    End Sub

    Private Sub replaceButton_Click(sender As Object, e As EventArgs) Handles replaceButton.Click
        Call MsgBox("noch nicht implementiert ...")
    End Sub

    Private Sub visElements_Click(sender As Object, e As EventArgs) Handles visElements.Click

        Call visualizeElements()

    End Sub
End Class