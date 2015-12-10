Imports ProjectBoardDefinitions

Public Class frmWbListInfo

    Friend phasesChecked As Boolean
    Friend calledFromStdList As Boolean

    ''' <summary>
    ''' löscht die selektierten Synonyme aus dem Wörterbuch  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deleteButton_Click(sender As Object, e As EventArgs) Handles deleteButton.Click

        Dim todoList As New Collection

        For i As Integer = 1 To ergebnisListe.SelectedItems.Count

            Dim synonym As String = ergebnisListe.SelectedItems.Item(i - 1)
            todoList.Add(synonym)
            If phasesChecked Then
                
                phaseMappings.removeSynonym(synonym)
                If Not missingPhaseDefinitions.Contains(synonym) Then

                    Dim phDef As New clsPhasenDefinition
                    With phDef
                        .name = synonym
                    End With

                    missingPhaseDefinitions.Add(phDef)

                End If
            Else

                milestoneMappings.removeSynonym(synonym)

                If Not missingMilestoneDefinitions.Contains(synonym) Then

                    Dim msDef As New clsMeilensteinDefinition
                    With msDef
                        .name = synonym
                    End With

                    missingMilestoneDefinitions.Add(msDef)

                End If

            End If
        Next

        ' jetzt aus der Listbox löschen 
        For i As Integer = 1 To todoList.Count

            ergebnisListe.Items.Remove(CStr(todoList.Item(i)))

        Next

    End Sub

    Private Sub frmWbListInfo_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        frmCoord(PTfrm.listInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.listInfo, PTpinfo.left) = Me.Left
    End Sub


    Private Sub frmWbListInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim anzahlSyn As Integer
        Dim syn As String
        Dim elemName As String = elementName.Text

        If frmCoord(PTfrm.listInfo, PTpinfo.top) = 0 And _
            frmCoord(PTfrm.listInfo, PTpinfo.left) = 0 Then
            Me.Top = 50
            Me.Left = 50
        Else
            Me.Top = frmCoord(PTfrm.listInfo, PTpinfo.top)
            Me.Left = frmCoord(PTfrm.listInfo, PTpinfo.left)
        End If

        ergebnisListe.Items.Clear()

        If calledFromStdList Then

            ' Remove Button sichtbar werden lassen 
            deleteButton.Visible = True

            ' headertext 
            headerText.Text = "dieses Element hat folgende Synonyme:"

            Me.Text = "Synonym Definitionen"


            If phasesChecked Then
                anzahlSyn = phaseMappings.countSynonyms(elemName)
                For i = 1 To anzahlSyn
                    syn = phaseMappings.getSynonymsOf(elemName, i)
                    ergebnisListe.Items.Add(syn)
                Next
            Else
                anzahlSyn = milestoneMappings.countSynonyms(elemName)
                For i = 1 To anzahlSyn
                    syn = milestoneMappings.getSynonymsOf(elemName, i)
                    ergebnisListe.Items.Add(syn)
                Next
            End If


        Else

            ' Remove Button unsichtbar werden lassen 
            deleteButton.Visible = False

            ' headertext 
            headerText.Text = "dieses Element kommt in folgenden Projekten vor: "

            If phasesChecked Then

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If Not IsNothing(kvp.Value.getPhase(elemName)) Then
                        ergebnisListe.Items.Add(kvp.Value.name)
                    End If

                Next

            Else

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If Not IsNothing(kvp.Value.getMilestone(elemName)) Then
                        ergebnisListe.Items.Add(kvp.Value.name)
                    End If

                Next

            End If
        End If

    End Sub

    Private Sub ergebnisListe_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ergebnisListe.SelectedIndexChanged

    End Sub
End Class