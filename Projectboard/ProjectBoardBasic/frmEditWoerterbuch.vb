Imports ProjectBoardDefinitions

Public Class frmEditWoerterbuch

    
    ' wird benutzt , um beim Listen Aktualisieren eine Aktion zu triggern oder eben nicht 
    Private eventsShouldFire As Boolean = True
    Private listOfSummaryTasks As New Collection
    Dim somethingChanged As Boolean = False

    Private Sub frmEditWoerterbuch_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Me.TopMost = False

        If somethingChanged Then

            If MsgBox(Prompt:="Bitte beachten:" & vbLf & "um die neuen Regeln und Definitionen anzuwenden," & vbLf & "müssen die Projekte neu importiert werden !" & vbLf & vbLf & _
                      "Sollen die Änderungen im Customization File gespeichert werden?", _
                      Buttons:=MsgBoxStyle.YesNo, _
                      Title:=" ") = MsgBoxResult.Yes Then

                Call awinWritePhaseMilestoneDefinitions(True)

                'Call MsgBox("Bitte beachten:" & vbLf & "um die neuen Regeln und Definitionen anzuwenden," & vbLf & "müssen die Projekte neu importiert werden !")
            End If

        End If


    End Sub
    

    ''' <summary>
    ''' triggert das Wegschreiben der Phasen-, Meilenstein-Definitionen und der 
    ''' Phasen- und Meilenstein Mappings  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmEditWoerterbuch_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        


    End Sub


    ''' <summary>
    ''' wird beim Laden des Formuars durchlaufen 
    ''' normalerweise wird gestartet mit Radio-Button Phases aktiv, 
    ''' ausser es gibt keine missing PhaseDefinitions, aber es gibt MissingMilestoneDefinitions 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmEditWoerterbuch_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        showOnlySummaryTasks.Checked = False
        ToolStripStatusLabel1.Text = ""

        ' jetzt Radio-Button setzen 
        If missingPhaseDefinitions.Count = 0 And missingMilestoneDefinitions.Count > 0 Then
            Me.rdbListShowsMilestones.Checked = True
        Else
            Me.rdbListShowsPhases.Checked = True
        End If


    End Sub

    ''' <summary>
    ''' wird getriggert, wenn Radiobutton Phasen aktiv/inaktiv  gesetzt wird 
    ''' je nachdem werden die Listen standardlist und unknownlist mit den Werten aus PhaseDefinitions, missingPhaseDefinitions bzw 
    ''' MilestoneDefinitions, missingMilestonedefinitions besetzt   
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbListShowsPhases_CheckedChanged(sender As Object, e As EventArgs) Handles rdbListShowsPhases.CheckedChanged

        showOnlySummaryTasks.Visible = True

        If listOfSummaryTasks.Count = 0 Then
            listOfSummaryTasks = buildListOfSummaryTasks()
        End If


        Dim tmpPhDef As clsPhasenDefinition
        Dim tmpMsDef As clsMeilensteinDefinition

        If rdbListShowsPhases.Checked Then

            unknownList.Items.Clear()
            standardList.Items.Clear()

            Call buildUnknownPhaseList()
           

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

            Call buildUnknownMilestoneList()

            

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

    ''' <summary>
    ''' wird getriggert, sobald sich der Filter zum Feld Unknown ändert 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub filterUnknown_TextChanged(sender As Object, e As EventArgs) Handles filterUnknown.TextChanged

        Dim suchstr As String = filterUnknown.Text
        Dim tmpName As String

        eventsShouldFire = False

        unknownList.Items.Clear()

        If rdbListShowsPhases.Checked Then

            Call buildUnknownPhaseList()

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

        eventsShouldFire = True

    End Sub

    ''' <summary>
    ''' wird getriggert, sobald sich der Text im Filter zu Standard ändert 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' löscht die Liste der unbekannten Bezeichnungen 
    ''' wenn in der Liste der Standard-Bezeichnungen 1 selektiert ist, wird das auch in das Feld EditUnknowItem geschrieben 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub clearUnknownList_Click(sender As Object, e As EventArgs) Handles clearUnknownList.Click

        unknownList.SelectedItems.Clear()
        If standardList.SelectedItems.Count = 1 Then
            editUnknownItem.Text = CStr(standardList.SelectedItem)
        End If

    End Sub


    ''' <summary>
    ''' löscht die Liste der Standard-Bezeichnungen 
    ''' wenn in der Liste der Unknown-Bezeichnungen 1 selektiert ist, wird das auch in das Feld EditUnknowItem geschrieben
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub clearStandardList_Click(sender As Object, e As EventArgs) Handles clearStandardList.Click

        standardList.SelectedItems.Clear()
        If unknownList.SelectedItems.Count = 1 Then
            editUnknownItem.Text = CStr(unknownList.SelectedItem)
        End If

    End Sub

    ''' <summary>
    ''' wird aufgerufen, wenn auf ein Item in der UnknowList ein Doppelklick gemacht wird 
    ''' zeigt die Liste der Projekte an, wo das Element vorkommt ...
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub unknownList_MouseDoubleClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles unknownList.MouseDoubleClick

        If unknownList.SelectedItems.Count = 1 Then

            Dim infoFrm As New frmWbListInfo
            infoFrm.elementName.Text = CStr(unknownList.SelectedItem)
            infoFrm.phasesChecked = rdbListShowsPhases.Checked
            infoFrm.calledFromStdList = False
            infoFrm.ShowDialog()

        Else
            Call MsgBox("Info kann nur für ein Element gezeigt werden ... ")
        End If

    End Sub

    ''' <summary>
    ''' wird aufgerufen, sobald sich was in der Selektion der unknownlist verändert
    ''' wenn nur ein Item selektiert ist, und in der anderen Liste nichts, wird das Eingabe Feld unten auf diesen Wert gesetzt 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub unknownList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles unknownList.SelectedIndexChanged

        If eventsShouldFire Then
            If unknownList.SelectedItems.Count = 1 And standardList.SelectedItems.Count = 0 Then
                editUnknownItem.Text = unknownList.SelectedItem.ToString
            Else
                editUnknownItem.Text = ""
            End If
        End If

        ToolStripStatusLabel1.Text = ""

    End Sub

    ''' <summary>
    ''' wird beim Doppelklick auf ein Element der Standardliste aufgerufen
    ''' zeigt die Liste der definierten Synonyme an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub standardList_MouseDoubleClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles standardList.MouseDoubleClick

        If standardList.SelectedItems.Count = 1 Then

            Dim infoFrm As New frmWbListInfo
            Dim anzahlElem As Integer
            If rdbListShowsPhases.Checked Then
                anzahlElem = missingPhaseDefinitions.Count
            Else
                anzahlElem = missingMilestoneDefinitions.Count
            End If

            infoFrm.elementName.Text = CStr(standardList.SelectedItem)
            infoFrm.phasesChecked = rdbListShowsPhases.Checked
            infoFrm.calledFromStdList = True
            infoFrm.ShowDialog()

            If rdbListShowsPhases.Checked Then
                If anzahlElem <> missingPhaseDefinitions.Count Then
                    eventsShouldFire = False
                    unknownList.Items.Clear()
                    Call buildUnknownPhaseList()
                    eventsShouldFire = True
                End If
            Else
                If anzahlElem <> missingMilestoneDefinitions.Count Then
                    eventsShouldFire = False
                    unknownList.Items.Clear()
                    Call buildUnknownMilestoneList()
                    eventsShouldFire = True
                End If
            End If
        Else
            Call MsgBox("Info kann nur für ein Element gezeigt werden ... ")
        End If

    End Sub

    ''' <summary>
    ''' wird aufgerufen, sobald sich was in der Selektion der standardList verändert
    ''' wenn nur ein Item selektiert ist, und in der unknownList nichts, wird das Eingabe Feld unten auf diesen Wert gesetzt 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub standardList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles standardList.SelectedIndexChanged

        If eventsShouldFire Then
            If standardList.SelectedItems.Count = 1 And unknownList.SelectedItems.Count = 0 Then
                editUnknownItem.Text = standardList.SelectedItem.ToString
            Else
                editUnknownItem.Text = ""
            End If

        End If

        ToolStripStatusLabel1.Text = ""

    End Sub


    ''' <summary>
    ''' visualisiert die selektierten Elemente aus der Unknown-List bzw. aus der Standard-List auf der Multiprojekt Tafel 
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
    ''' macht eine bisher unbekannte Bezeichnung zum Standard , also ergänzt ..Definitions und löscht aus ..missingDefinitions
    ''' der String kann zuvor noch editiert werden; in diesem Fall wird nicht aus missingDefinitions gelöscht; 
    ''' Bedingung: es dürfen mehrere Elemente aus der unknownList selektiert sein, aber kein einziges aus der StandardList
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub makeItemToBeStandard()

        Dim itemText As String
        Dim editText As String
        Dim anzahlElements As Integer = unknownList.SelectedItems.Count

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
                                standardList.Items.Add(itemText)

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

        editUnknownItem.Text = ""
        If anzahlElements = 1 Then
            ToolStripStatusLabel1.Text = "ok, 1 Element als Standard-Name aufgenommen"
        ElseIf anzahlElements > 1 Then
            ToolStripStatusLabel1.Text = "ok, " & anzahlElements & " Elemente als Standard-Name aufgenommen"
        End If



        eventsShouldFire = True

    End Sub

    

    ''' <summary>
    ''' ruft die Methode auf, um ein Item aus der unbekannten Liste zu einem Standard-Item zu machen 
    ''' die entsprechenden Definitionen werden dabei aus der missingPhaseDefinitions übernommen  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub setItemToBeKnown_Click(sender As Object, e As EventArgs) Handles setItemToBeKnown.Click

        somethingChanged = True
        Call makeItemToBeStandard()



    End Sub


    ''' <summary>
    ''' baut in Abhängigkeit vom Filter und dem Status von ShowSummaryTasksOnly die Liste der Unknown-Elemente auf 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub buildUnknownPhaseList()

        Dim suchstr As String = filterUnknown.Text
        Dim tmpName As String



        For i = 1 To missingPhaseDefinitions.Count
            tmpName = missingPhaseDefinitions.getPhaseDef(i).name
            If filterUnknown.Text = "" Then
                If showOnlySummaryTasks.Checked Then
                    If listOfSummaryTasks.Contains(tmpName) Then
                        unknownList.Items.Add(tmpName)
                    End If
                Else
                    unknownList.Items.Add(tmpName)
                End If

            Else
                If tmpName.Contains(suchstr) Then
                    If showOnlySummaryTasks.Checked Then
                        If listOfSummaryTasks.Contains(tmpName) Then
                            unknownList.Items.Add(tmpName)
                        End If
                    Else
                        unknownList.Items.Add(tmpName)
                    End If
                End If
            End If
        Next

        
    End Sub

    ''' <summary>
    ''' baut in Abhängigkeit von Filter aus der Missing-Definitions die Meilenstein Liste auf 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub buildUnknownMilestoneList()
        ' Meilensteine
        Dim tmpName As String
        Dim suchstr As String = filterUnknown.Text

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


    End Sub

    ''' <summary>
    ''' erzeugt aus den aktuell geladenen Projekten die Liste aller Phasen, die Summary Tasks sind 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function buildListOfSummaryTasks() As Collection

        Dim lastPhaseIndex As Integer
        Dim tmpCollection As New Collection
        Dim nodeItem As clsHierarchyNode

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            lastPhaseIndex = kvp.Value.hierarchy.getIndexOf1stMilestone - 1
            If lastPhaseIndex < 0 Then
                ' es gibt keine Meilensteine, sondern nur Phasen 
                lastPhaseIndex = kvp.Value.hierarchy.count
            End If

            For i As Integer = 1 To lastPhaseIndex
                nodeItem = kvp.Value.hierarchy.nodeItem(i)
                If nodeItem.childCount > 0 Then
                    If Not tmpCollection.Contains(nodeItem.elemName) Then
                        tmpCollection.Add(nodeItem.elemName, nodeItem.elemName)
                    End If
                End If
            Next


        Next

        buildListOfSummaryTasks = tmpCollection

    End Function

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
        Dim anzahlElements As Integer = standardList.SelectedItems.Count

        eventsShouldFire = False
        somethingChanged = True

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
                        Call MsgBox(itemText & " hat " & anzEintraege & " Einträge im Wörterbuch. Trotzdem fortfahren?", MsgBoxStyle.OkCancel)
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
                        Call MsgBox(itemText & " hat " & anzEintraege & " Einträge im Wörterbuch. Trotzdem fortfahren?", MsgBoxStyle.OkCancel)
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

        editUnknownItem.Text = ""

        If anzahlElements = 1 Then
            ToolStripStatusLabel1.Text = "ok, 1 Element aus Standard-Liste entfernt"
        ElseIf anzahlElements > 1 Then
            ToolStripStatusLabel1.Text = "ok, " & anzahlElements & " Elemente aus Standard-Liste entfernt"
        End If



        eventsShouldFire = True
        

    End Sub

    ''' <summary>
    ''' fügt den oder die selektierten Namen in die Liste der zu ignorierenden Elemente ein
    ''' Vorbedingung: kein Element aus der Standard-Liste darf selektiert sein 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ignoreButton_Click(sender As Object, e As EventArgs) Handles ignoreButton.Click

        Dim anzahlElements As Integer = unknownList.SelectedItems.Count

        eventsShouldFire = False

        somethingChanged = True

        If standardList.SelectedItems.Count > 0 Then
            Call MsgBox("in der Liste der bekannten Namen darf kein Eintrag selektiert sein ...")
        Else
            If unknownList.SelectedItems.Count > 0 Then

                Dim todoList As New Collection

                If rdbListShowsPhases.Checked Then

                    For Each Obj As Object In unknownList.SelectedItems
                        Try
                            Dim itemName As String = CStr(Obj)
                            phaseMappings.addIgnoreName(itemName)
                            missingPhaseDefinitions.remove(itemName)
                            todoList.Add(itemName)
                            
                        Catch ex As Exception

                        End Try
                    Next

                Else

                    For Each Obj As Object In unknownList.SelectedItems
                        Try
                            Dim itemName As String = CStr(Obj)
                            milestoneMappings.addIgnoreName(itemName)
                            missingMilestoneDefinitions.remove(itemName)
                            todoList.Add(itemName)
                        Catch ex As Exception

                        End Try
                    Next

                End If

                ' jetzt müssen aus der unknownList alle Einträge der todoList raus 
                For i As Integer = 1 To todoList.Count
                    Dim itemName As String = CStr(todoList.Item(i))
                    unknownList.Items.Remove(itemName)
                Next

                todoList.Clear()

            Else
                Call MsgBox("wenigstens ein Element aus der Liste der unbekannten Bezeichnungen selektieren ...")
            End If
        End If

        If anzahlElements > 0 Then

        End If

        editUnknownItem.Text = ""
        If anzahlElements = 1 Then
            ToolStripStatusLabel1.Text = "ok, 1 Element wird zukünftig ignoriert ..."
        ElseIf anzahlElements > 1 Then
            ToolStripStatusLabel1.Text = "ok, " & anzahlElements & " Element(e werden zukünftig ignoriert ..."
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
        Dim anzahlElements As Integer = unknownList.SelectedItems.Count

        eventsShouldFire = False

        somethingChanged = True

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

        editUnknownItem.Text = ""

        If anzahlElements = 1 Then
            ToolStripStatusLabel1.Text = "ok, 1 Regel wurde im Wörterbuch aufgenommen  ..."
        ElseIf anzahlElements > 1 Then
            ToolStripStatusLabel1.Text = "ok, " & anzahlElements & " Regeln wurden im Wörterbuch aufgenommen  ..."
        End If

        eventsShouldFire = True
        
    End Sub

    ''' <summary>
    ''' ersetzt eine Standardbezeichnung durch eine andere; dabei werden auch alle Abbildungsregeln entsprechend auf den neuen Wert gesetzt;
    ''' nimmt auch eine neue Abbildungsregel auf : alter Wert -> neuer Wert 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub replaceButton_Click(sender As Object, e As EventArgs) Handles replaceButton.Click

        eventsShouldFire = False

        somethingChanged = True

        ' erst müssen alle Wörterbuch Einträge ersetzt werden, die bisher auf diesen Standard-Namen abbilden
        Dim newStdName As String = editUnknownItem.Text
        Dim oldStdName As String = CStr(standardList.SelectedItem)

        If newStdName = oldStdName Then
            ' nichts machen , sind identisch 
            Call MsgBox("beide Namen sind identisch - keine Ersetzung vorgenommen")
        Else
            If rdbListShowsPhases.Checked Then
                ' Phasen
                phaseMappings.replaceInSynonyms(oldStdName, newStdName)
                Try
                    phaseMappings.addSynonym(oldStdName, newStdName)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try

                ' und jetzt muss der Eintrag in phaseDefinitions noch geändert werden
                Dim phaseDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(oldStdName)
                phaseDef.name = newStdName
                PhaseDefinitions.remove(oldStdName)
                PhaseDefinitions.Add(phaseDef)

                ' löschen in der StandardListe
                standardList.Items.Remove(oldStdName)
                standardList.Items.Add(newStdName)

                ' wenn vorhanden in MissingPhaseDefinitions: löschen , auch in der Unknownlist
                If missingPhaseDefinitions.Contains(newStdName) Then
                    missingPhaseDefinitions.remove(newStdName)
                    unknownList.Items.Remove(newStdName)
                End If


            Else
                ' Meilensteine
                milestoneMappings.replaceInSynonyms(oldStdName, newStdName)
                Try
                    milestoneMappings.addSynonym(oldStdName, newStdName)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try

                ' und jetzt muss der Eintrag in milestoneDefinitions noch geändert werden 

                Dim milestoneDef As clsMeilensteinDefinition = MilestoneDefinitions.getMilestoneDef(oldStdName)
                milestoneDef.name = newStdName
                MilestoneDefinitions.remove(oldStdName)
                MilestoneDefinitions.Add(milestoneDef)

                ' löschen in der StandardListe
                standardList.Items.Remove(oldStdName)
                standardList.Items.Add(newStdName)

                ' wenn vorhanden in MissingPhaseDefinitions: löschen , auch in der Unknownlist
                If missingMilestoneDefinitions.Contains(newStdName) Then
                    missingMilestoneDefinitions.remove(newStdName)
                    unknownList.Items.Remove(newStdName)
                End If
            End If
        End If

        editUnknownItem.Text = ""
        ToolStripStatusLabel1.Text = "ok, " & oldStdName & " wurde durch " & newStdName & " ersetzt ..."

        eventsShouldFire = True

    End Sub

    ''' <summary>
    ''' visualisiert die selektierten Elemente aus der Standard- und Unknownlist auf der Multiprojekt-Tafel
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub visElements_Click(sender As Object, e As EventArgs) Handles visElements.Click

        Call awinDeSelect()

        Call visualizeElements()

    End Sub

    Private Sub showOnlySummaryTasks_CheckedChanged(sender As Object, e As EventArgs) Handles showOnlySummaryTasks.CheckedChanged


        eventsShouldFire = False

        unknownList.Items.Clear()

        Call buildUnknownPhaseList()

        editUnknownItem.Text = ""

        eventsShouldFire = True

    End Sub

    Private Sub rdbListShowsMilestones_CheckedChanged(sender As Object, e As EventArgs) Handles rdbListShowsMilestones.CheckedChanged

        showOnlySummaryTasks.Visible = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Call awinWritePhaseMilestoneDefinitions(True)
        somethingChanged = False

        ToolStripStatusLabel1.Text = "ok, Änderungen wurden gespeichert ..."

        Call MsgBox("Bitte beachten:" & vbLf & "um die neuen Regeln und Definitionen anzuwenden," & vbLf & "müssen die Projekte neu importiert werden !")

    End Sub
End Class