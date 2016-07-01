Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports MongoDbAccess
Imports System.Windows.Forms

''' <summary>
''' wird verwendet um Portfolios zu definieren, Varianten zu aktivieren, Projekte und Varianten zu laden, zu aktivieren und zu löschen 
''' </summary>
''' <remarks></remarks>
Public Class frmProjPortfolioAdmin

    Private aktuelleGesamtListe As New clsProjekteAlle
    Private projektHistorien As New clsProjektDBInfos
    Private stopRecursion As Boolean = False
    Private constellationName As String = ""


    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    ' wird an der aufrufenden Stelle gesetzt; steuert, was mit den ausgewählten ELementen geschieht
    Friend aKtionskennung As Integer

    Private Sub frmDefineEditPortfolio_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left
        projektHistorien.clear()

    End Sub

    Private Sub frmDefineEditPortfolio_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If frmCoord(PTfrm.eingabeProj, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.top))
        End If

        If frmCoord(PTfrm.eingabeProj, PTpinfo.left) > 0 Then
            Me.Left = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.left))
        End If


        ' alle definierten Filter in ComboBox anzeigen
        ' die Filter einlesen 
        Call frmHryNameReadFilterVorlagen(PTmenue.filterdefinieren, filterDropbox)
        For Each kvp As KeyValuePair(Of String, clsFilter) In filterDefinitions.Liste
            filterDropbox.Items.Add(kvp.Key)
        Next


        stopRecursion = True
        Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, aKtionskennung, _
                           False)
        stopRecursion = False

        If aktuelleGesamtListe.liste.Count < 1 Then
            DialogResult = Windows.Forms.DialogResult.OK
        End If

    End Sub

    Private Sub TreeViewProjekte_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterCheck

        Dim node As TreeNode
        Dim schluessel As String = ""
        'Dim selCollection As SortedList(Of Date, String)
        'Dim timeStamp As Date
        Dim treeLevel As Integer
        Dim i As Integer, j As Integer
        Dim childNode As TreeNode
        Dim parentNode As TreeNode

        ' Andernfalls wird die Check Routine endlos aufgerufen ...
        If stopRecursion Then
            Exit Sub
        End If

        node = e.Node
        treeLevel = node.Level



        ' hier wird jetzt sichergestellt, daß nur die nach der aktuellen Aktion gültigen Checks gesetzt werden können

        If aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.loadPV Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                        For j = 1 To childNode.Nodes.Count
                            childNode.Nodes.Item(j - 1).Checked = node.Checked
                        Next
                    Next

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' nach unten: das Gleiche 
                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                    Next
                    ' nach oben 

                    If node.Checked = False Then
                        node.Parent.Checked = False
                    End If

                    ' wenn mit diesem Knoten jetzt alle gesetzt sind, soll auch parent wieder gesetzt werden 
                    If node.Checked = True Then
                        parentNode = node.Parent
                        Dim allchecked As Boolean = True
                        For i = 1 To parentNode.Nodes.Count
                            allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                        Next
                        If allchecked Then
                            parentNode.Checked = True
                        End If
                    End If

                Case 2 ' Snapshot ist selektiert / nicht selektiert 
                    If node.Checked = False Then
                        node.Parent.Checked = False
                        parentNode = node.Parent
                        parentNode.Parent.Checked = False
                    End If

                    If node.Checked = True Then
                        parentNode = node.Parent
                        Dim allchecked As Boolean = True
                        For i = 1 To parentNode.Nodes.Count
                            allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                        Next
                        If allchecked Then
                            ' jetzt wird bewusst Rekursion angestossen, damit das nach oben weitergeht ...
                            stopRecursion = False
                            parentNode.Checked = True
                        End If
                    End If

            End Select

            stopRecursion = False


        ElseIf aKtionskennung = PTTvActions.delFromSession Or _
              aKtionskennung = PTTvActions.deleteV Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                        For j = 1 To childNode.Nodes.Count
                            childNode.Nodes.Item(j - 1).Checked = node.Checked
                        Next
                    Next

                Case 1 ' Variante ist selektiert / nicht selektiert

                    ' nach unten: das Gleiche 
                    For i = 1 To node.Nodes.Count
                        childNode = node.Nodes.Item(i - 1)
                        childNode.Checked = node.Checked
                    Next
                    ' nach oben 

                    If node.Checked = False Then
                        node.Parent.Checked = False
                    End If

                    ' wenn mit diesem Knoten jetzt alle gesetzt sind, soll auch parent wieder gesetzt werden 
                    If node.Checked = True Then
                        parentNode = node.Parent
                        Dim allchecked As Boolean = True
                        For i = 1 To parentNode.Nodes.Count
                            allchecked = allchecked And parentNode.Nodes.Item(i - 1).Checked
                        Next
                        If allchecked Then
                            parentNode.Checked = True
                        End If
                    End If


            End Select

            stopRecursion = False
        ElseIf aKtionskennung = PTTvActions.activateV Or _
               aKtionskennung = PTTvActions.definePortfolioDB Or _
               aKtionskennung = PTTvActions.definePortfolioSE Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    ' bei Aktivieren kann man Projekt nicht selektieren 
                    node.Checked = False

                Case 1 ' Variante ist selektiert / nicht selektiert


                    Dim projektNode As TreeNode = node.Parent
                    Dim selectedVariantName As String = node.Text
                    Dim pName As String = projektNode.Text

                    ' es kann immer nur eine Variante selektiert sein; wenn die bisher aktive de-selektiert wird, 
                    ' wird Standard auf checked gesetzt 

                    If node.Checked = True Then

                        ' alle anderen Varianten auf Unchecked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Text <> selectedVariantName Then
                                projektNode.Nodes.Item(i).Checked = False
                            End If
                        Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        selectedVariantName = getVariantNameOf(node.Text)



                    Else

                        ' die Standard Variante auf Checked setzen 
                        For i = 0 To projektNode.Nodes.Count - 1
                            If projektNode.Nodes.Item(i).Text = "()" Then
                                projektNode.Nodes.Item(i).Checked = True
                            End If
                        Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        ' aber nur, wenn es nicht vorher schon die leere Variante war 

                        selectedVariantName = ""

                    End If

                    If aKtionskennung = PTTvActions.activateV Then
                        ' jetzt die Variante aktivieren 
                        Call replaceProjectVariant(pName, selectedVariantName, True, True, 0)
                        Call awinNeuZeichnenDiagramme(2)
                    End If



            End Select

            stopRecursion = False

        End If


    End Sub

    Private Sub TreeViewProjekte_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterSelect

        Dim node As TreeNode = e.Node

        'If node.IsSelected Then
        '    node.Expand()
        'End If

    End Sub

    Private Sub TreeViewProjekte_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeViewProjekte.BeforeExpand

        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim node As New TreeNode
        Dim nodeVariant As New TreeNode
        Dim nodeTimeStamp As New TreeNode
        Dim projName As String = ""
        Dim variantName As String = ""
        Dim hliste As SortedList(Of Date, String)
        Dim nodeLevel As Integer
        Dim variantListe As Collection
        Dim hproj As New clsProjekt
        Dim key As String



        node = e.Node
        nodeLevel = node.Level

        If nodeLevel = 0 Then
            projName = node.Text

            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If node.Tag = "P" Then
                ' Inhalte der Sub-Nodes müssen neu aufgebaut werden 
                variantListe = aktuelleGesamtListe.getVariantNames(projName, True)

                ' hproj wird benötigt, um herauszufinden, welche Variante gerade aktiv ist
                If aKtionskennung = PTTvActions.activateV Then
                    hproj = ShowProjekte.getProject(projName)
                End If


                ' Löschen von Platzhalter
                node.Nodes.Clear()

                ' Eintragen der zum Projekt gehörenden Varianten
                For Each variantName In variantListe
                    nodeVariant = node.Nodes.Add(CType(variantName, String))

                    ' jetzt muss gecheckt werden , ob es sich um das Aktivieren handelt oder nicht
                    If aKtionskennung = PTTvActions.activateV Then
                        stopRecursion = True
                        If getVariantNameOf(variantName) = hproj.variantName Then
                            nodeVariant.Checked = True
                        Else
                            nodeVariant.Checked = False
                        End If
                        stopRecursion = False

                    ElseIf aKtionskennung = PTTvActions.loadPV Then

                        key = calcProjektKey(pName:=projName, variantName:=variantName)

                        stopRecursion = True
                        ' soll gesetzt sein, wenn es entweder bereits geladen ist oder aber sowieso alle geladen werden sollen
                        If AlleProjekte.Containskey(key) Or node.Checked = True Then
                            nodeVariant.Checked = True
                        Else
                            nodeVariant.Checked = False
                        End If
                        stopRecursion = False

                    Else
                        nodeVariant.Checked = node.Checked
                    End If



                    If aKtionskennung = PTTvActions.delFromDB Or _
                        aKtionskennung = PTTvActions.loadPVS Then
                        ' Einfügen eines Platzhalters macht nur Sinn bei Snapshots löschen bzw. Snapshots laden 

                        nodeVariant.Tag = "P"
                        nodeVariant.Nodes.Add("()")
                    Else
                        nodeVariant.Tag = "X"
                    End If

                Next

                node.Tag = "X"



            End If



        ElseIf nodeLevel = 1 And _
            (aKtionskennung = PTTvActions.delFromDB Or aKtionskennung = PTTvActions.loadPVS) Then


            If node.Tag = "P" Then

                node.Tag = "X"
                projName = node.Parent.Text
                variantName = getVariantNameOf(node.Text)

                hliste = projektHistorien.getTimeStamps(calcProjektKey(projName, variantName))

                If hliste.Count = 0 Then

                    If Not noDB Then

                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        If request.pingMongoDb() Then
                        Else
                            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                        End If

                        ' Lesen der TimeStamp Snapshots für ProjNAme, variantName 
                        Try
                            If Not projekthistorie Is Nothing Then
                                projekthistorie.clear()
                            Else
                                projekthistorie = New clsProjektHistorie
                            End If

                            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=projName, variantName:=variantName, _
                                                                             storedEarliest:=Date.MinValue, storedLatest:=Date.Now)

                        Catch ex As Exception
                            projekthistorie.clear()
                        End Try

                    End If

                    If projekthistorie.Count > 0 Then

                        projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                        node.Nodes.Clear()  ' Löschen von Platzhalter

                        ' Aufbau der Listen 
                        projektHistorien.Add(projekthistorie)


                        ' Eintragen der zur Projekt-Variante gehörenden TimeStamps
                        For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                            nodeTimeStamp = node.Nodes.Add(CType(kvp1.Value.timeStamp, String))
                            nodeTimeStamp.Checked = node.Checked
                        Next kvp1


                    Else

                        If projekthistorie.Count = 0 Then
                            ' keine ProjektHistorie vorhanden
                            projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                            node.Nodes.Clear()  ' Löschen von Platzhalter
                        End If
                    End If




                End If

            End If


        End If





    End Sub

    ''' <summary>
    ''' liefert den Namen der Variante zurück, bereinigt um die öffnende und schließende Klammer
    ''' </summary>
    ''' <param name="nodeText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getVariantNameOf(ByVal nodeText As String) As String
        Dim tmpstr() As String
        Dim vName As String = ""

        tmpstr = nodeText.Split(New Char() {CChar("("), CChar(")")}, 3)
        If tmpstr.Length = 3 Then
            vName = tmpstr(1)
        End If

        getVariantNameOf = vName

    End Function

    ''' <summary>
    ''' wird bei Auslösen des "Aktionsbuttons" ausgeführt; 
    ''' in Abhängigkeit von Aktionskennung 
    ''' dieser Button kann im Fall activate Variante gar nicht aktiviert werden, weil unsichtbar
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim projektNode As TreeNode, variantNode As TreeNode, timeStampNode As TreeNode
        Dim anzahlProjekte As Integer
        Dim anzahlVarianten As Integer
        Dim anzahlTimeStamps As Integer
        Dim pname As String, variantName As String, timestamp As Date
        'Dim hproj As clsProjekt
        Dim portfolioZeile As Integer = 2

        ' ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        ' ''Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

        Dim p As Integer, v As Integer, t As Integer

        ' ''ur:10.08.2015 '' ''
        ' '' '' '' Aktivieren von Varianten erfordert überhaupt keinen Button; deswegen ist das jetzt hier nicht abgefragt 
        ' '' '' ''
        '' '' ''If aKtionskennung = PTTvActions.definePortfolioSE Or _
        '' '' ''    aKtionskennung = PTTvActions.definePortfolioDB Then
        '' '' ''    '
        '' '' ''    ' Portfolios definieren 
        '' '' ''    '
        '' '' ''    ' prüfen, ob diese Constellation bereits existiert ..


        '' '' ''    If IsNothing(portfolioName.SelectedItem) Then
        '' '' ''        constellationName = portfolioName.Text
        '' '' ''    Else
        '' '' ''        constellationName = portfolioName.SelectedItem.ToString
        '' '' ''    End If

        '' '' ''    If constellationName.Length = 0 Then
        '' '' ''        Call MsgBox("bitte einen Namen angeben")
        '' '' ''        Exit Sub
        '' '' ''    End If

        '' '' ''    If projectConstellations.Contains(constellationName) Then

        '' '' ''        Try
        '' '' ''            projectConstellations.Remove(constellationName)
        '' '' ''        Catch ex As Exception

        '' '' ''        End Try

        '' '' ''    End If

        '' '' ''    Dim newC As New clsConstellation
        '' '' ''    With newC
        '' '' ''        .constellationName = constellationName
        '' '' ''    End With


        '' '' ''    With TreeViewProjekte
        '' '' ''        anzahlProjekte = .Nodes.Count

        '' '' ''        For p = 1 To anzahlProjekte

        '' '' ''            projektNode = .Nodes.Item(p - 1)
        '' '' ''            pname = projektNode.Text
        '' '' ''            variantName = ""

        '' '' ''            If projektNode.Checked Then
        '' '' ''                ' das Projekt mit Variante "" in Konstellation eintragen

        '' '' ''                hproj = request.retrieveOneProjectfromDB(pname, variantName)

        '' '' ''                Dim newConstellationItem As New clsConstellationItem

        '' '' ''                With newConstellationItem
        '' '' ''                    .projectName = pname
        '' '' ''                    .show = True
        '' '' ''                    .Start = hproj.startDate
        '' '' ''                    .variantName = hproj.variantName
        '' '' ''                    .zeile = portfolioZeile
        '' '' ''                    portfolioZeile = portfolioZeile + 1
        '' '' ''                End With

        '' '' ''                newC.Add(newConstellationItem)


        '' '' ''                ' wenn es bereits ersetzt wurde, dann stimmt anzahlVarianten = ... 
        '' '' ''            ElseIf projektNode.Tag = "X" Then

        '' '' ''                anzahlVarianten = projektNode.Nodes.Count

        '' '' ''                For v = 1 To anzahlVarianten
        '' '' ''                    variantNode = projektNode.Nodes.Item(v - 1)
        '' '' ''                    variantName = getVariantNameOf(variantNode.Text)


        '' '' ''                    If variantNode.Checked Then

        '' '' ''                        hproj = request.retrieveOneProjectfromDB(pname, variantName)

        '' '' ''                        Dim newConstellationItem As New clsConstellationItem

        '' '' ''                        With newConstellationItem
        '' '' ''                            .projectName = pname
        '' '' ''                            .show = True
        '' '' ''                            .Start = hproj.startDate
        '' '' ''                            .variantName = hproj.variantName
        '' '' ''                            .zeile = portfolioZeile
        '' '' ''                            portfolioZeile = portfolioZeile + 1
        '' '' ''                        End With

        '' '' ''                        newC.Add(newConstellationItem)



        '' '' ''                    End If

        '' '' ''                Next
        '' '' ''            End If

        '' '' ''        Next


        '' '' ''        Try
        '' '' ''            projectConstellations.Add(newC)
        '' '' ''            Call MsgBox("Portfolio " & constellationName & " gespeichert ...")
        '' '' ''        Catch ex As Exception
        '' '' ''            Call MsgBox("Fehler bei Add projectConstellations in awinStoreConstellations")
        '' '' ''        End Try

        '' '' ''        ' Portfolio in die Datenbank speichern, falls Aktionskennung 
        '' '' ''        If aKtionskennung = PTTvActions.definePortfolioDB Then
        '' '' ''            If request.pingMongoDb() Then
        '' '' ''                If Not request.storeConstellationToDB(newC) Then
        '' '' ''                    Call MsgBox("Fehler beim Speichern der ProjektConstellation '" & newC.constellationName & "' in die Datenbank")
        '' '' ''                End If
        '' '' ''            Else
        '' '' ''                Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
        '' '' ''            End If
        '' '' ''        End If


        '' '' ''    End With

        If aKtionskennung = PTTvActions.delFromDB Or _
            aKtionskennung = PTTvActions.delFromSession Or _
            aKtionskennung = PTTvActions.deleteV Or _
            aKtionskennung = PTTvActions.loadPV Then

            ' alle anderen Aktionen wie Projekte aus Datenbank löschen , aus Session löschen, aus Datenbank laden  ... 
            With TreeViewProjekte
                anzahlProjekte = .Nodes.Count

                For p = 1 To anzahlProjekte

                    projektNode = .Nodes.Item(p - 1)
                    pname = projektNode.Text

                    If projektNode.Checked Then
                        ' Aktion auf allen Varianten und Timestamps 
                        ' Schleife über alle Varianten: 
                        ' lösche in Datenbank pname#vname

                        'anzahlVarianten = projektNode.Nodes.Count

                        Dim variantListe As Collection = aktuelleGesamtListe.getVariantNames(pname, True)
                        anzahlVarianten = variantListe.Count

                        If aKtionskennung = PTTvActions.delFromSession Then

                            Call awinDeleteProjectInSession(pName:=pname)

                        ElseIf aKtionskennung = PTTvActions.delFromDB Then


                                For v = 1 To anzahlVarianten

                                    'variantNode = projektNode.Nodes.Item(v - 1)
                                    'variantName = getVariantNameOf(variantNode.Text)
                                    variantName = getVariantNameOf(CStr(variantListe.Item(v)))
                                    Call deleteCompleteProjectVariant(pname, variantName, aKtionskennung)

                                Next
                            

                        ElseIf aKtionskennung = PTTvActions.loadPV Then

                           
                                For v = 1 To anzahlVarianten

                                    'variantNode = projektNode.Nodes.Item(v - 1)
                                    'variantName = getVariantNameOf(variantNode.Text)
                                    variantName = getVariantNameOf(CStr(variantListe.Item(v)))

                                    If v = 1 Then
                                        Call loadProjectfromDB(pname, variantName, True)
                                    Else
                                        Call loadProjectfromDB(pname, variantName, False)
                                    End If


                                Next



                        End If




                    ElseIf projektNode.Tag = "X" Then

                        anzahlVarianten = projektNode.Nodes.Count
                        Dim first As Boolean = True

                        For v = 1 To anzahlVarianten
                            variantNode = projektNode.Nodes.Item(v - 1)
                            variantName = getVariantNameOf(variantNode.Text)


                            If variantNode.Checked Then
                                ' Aktion auf allen Timestamps
                                ' lösche in Datenbank das Objekt mit DB-Namen pname#vname

                                If aKtionskennung = PTTvActions.delFromDB Or _
                                    aKtionskennung = PTTvActions.delFromSession Or _
                                    aKtionskennung = PTTvActions.deleteV Then
                                    Call deleteCompleteProjectVariant(pname, variantName, aKtionskennung)

                                ElseIf aKtionskennung = PTTvActions.loadPV Then

                                    Call loadProjectfromDB(pname, variantName, first)
                                    first = False

                                End If


                            ElseIf aKtionskennung = PTTvActions.delFromDB Or _
                                    aKtionskennung = PTTvActions.loadPVS Then

                                anzahlTimeStamps = variantNode.Nodes.Count
                                Dim firstTS As Boolean = True
                                For t = 1 To anzahlTimeStamps
                                    timeStampNode = variantNode.Nodes.Item(t - 1)

                                    If timeStampNode.Checked Then
                                        ' Aktion auf diesem timestamp

                                        timestamp = CType(timeStampNode.Text, Date)
                                        If aKtionskennung = PTTvActions.delFromDB Then
                                            Call deleteProjectVariantTimeStamp(pname, variantName, timestamp, firstTS)
                                        Else
                                            ' Aktion für LoadPVS : aber hier gibt es wahrscheinlich gar keinen OK-Button
                                        End If

                                    End If
                                Next
                            End If

                        Next
                    End If

                Next

                If aKtionskennung = PTTvActions.loadPV Or _
                    aKtionskennung = PTTvActions.delFromSession Then
                    Call awinNeuZeichnenDiagramme(2)
                End If

            End With

            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()

        Else
            Call MsgBox("nicht unterstützte Option in ProjPortfolio Admin Formular ...")
        End If



    End Sub

    ' 'ur: 11.08.2015: PT-103' ''Private Sub applyFilter_CheckedChanged(sender As Object, e As EventArgs)

    ' '' ''    ' der TreeView muss neu aufgebaut werden 

    ' '' ''    stopRecursion = True
    ' '' ''    Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, aKtionskennung, _
    ' '' ''                       Me.applyFilter.Checked)
    ' '' ''    stopRecursion = False


    ' '' ''End Sub

    ' ''ur: 11.08.2015: PT-103 '' ''' <summary>
    ' '' '' ''' ruft die Eingabe Maske zum Definieren des Filters auf 
    ' '' '' ''' </summary>
    ' '' '' ''' <param name="sender"></param>
    ' '' '' ''' <param name="e"></param>
    ' '' '' ''' <remarks></remarks>
    ' '' ''Private Sub defineFilter_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles defineFilter.LinkClicked

    ' '' ''    Call defineFilterDB()

    ' '' ''End Sub

    Private Sub filterDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles filterDropbox.SelectedIndexChanged

        Dim fName As String = filterDropbox.SelectedItem.ToString
        ' wird nicht benötigt: ur: 29.07.2015 Dim filter As clsFilter = filterDefinitions.retrieveFilter(fName)

        ' jetzt werden anhand des Filters "fName" die Collections gesetzt 
        Call retrieveSelections(fName, PTmenue.filterAuswahl, selectedBUs, selectedTyps, _
                                selectedPhases, selectedMilestones, _
                                selectedRoles, selectedCosts)
        ' und als "Last" gespeichert
        Call storeFilter("Last", PTmenue.filterAuswahl, selectedBUs, selectedTyps, _
                                selectedPhases, selectedMilestones, _
                                selectedRoles, selectedCosts, False)

        stopRecursion = True
        Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, aKtionskennung, _
                           True)
        stopRecursion = False

        '  Call MsgBox("in filterDropbox")


    End Sub
End Class