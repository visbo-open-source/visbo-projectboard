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

    Private filterIsActive As Boolean = False
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

        ' Maus auf Normalmodus zurücksetzen
        appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

    End Sub

    Private Sub frmDefineEditPortfolio_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If frmCoord(PTfrm.eingabeProj, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.top))
        End If

        If frmCoord(PTfrm.eingabeProj, PTpinfo.left) > 0 Then
            Me.Left = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.left))
        End If


        ' jetzt die vorkommenden Timestamps auslesen 
        ' aber nicht bei allen Aktionskennungen 

        If aKtionskennung = PTTvActions.chgInSession Or _
            aKtionskennung = PTTvActions.delFromSession Or _
            aKtionskennung = PTTvActions.deleteV Or _
            aKtionskennung = PTTvActions.activateV Then
            dropBoxTimeStamps.Visible = False
            lblStandvom.Visible = False
        Else

            Try
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                Dim tCollection As Collection = request.retrieveZeitstempelFromDB()
                'Dim heute As String = Date.Now.ToString

                dropBoxTimeStamps.Items.Clear()

                For k As Integer = 1 To tCollection.Count
                    Dim tmpDate As Date = CDate(tCollection.Item(k))
                    dropBoxTimeStamps.Items.Add(tmpDate)
                Next

            Catch ex As Exception

            End Try
            
            ' jetzt ist dropBoxTimeStamps.selecteditem = Nothing ..
        End If


        ' Maus auf Wartemodus setzen
        appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

        If aKtionskennung = PTTvActions.chgInSession Then
            Me.considerDependencies.Visible = True
            Me.considerDependencies.Checked = False
            Me.dropboxScenarioNames.Text = "Multiprojekt-Szenario"
            Me.OKButton.Text = "Szenario speichern"
            For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                If kvp.Key <> "Start" Then
                    dropboxScenarioNames.Items.Add(kvp.Key)
                End If
            Next
        Else
            ' nichts weiter tun 
            ' die Filter-NAmen müssen aktuell nicht ausgelesen werden 
        End If



        stopRecursion = True
        Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, aKtionskennung, _
                           False, Date.Now)
        stopRecursion = False

        If aktuelleGesamtListe.liste.Count < 1 Then
            DialogResult = Windows.Forms.DialogResult.OK
        End If

        ' Maus auf Normalmodus zurücksetzen
        appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault


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
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOf(projektNode.Nodes.Item(0).Text)
                        Else
                            ' darf eigentlich gar nicht vorkommen 
                            selectedVariantName = ""
                        End If

                        ' alt , wurde durch oberes ersetzt; Schwäche war: wenn kein Varianten-Name "" existiert 
                        'For i = 0 To projektNode.Nodes.Count - 1
                        '    If projektNode.Nodes.Item(i).Text = "()" Then
                        '        projektNode.Nodes.Item(i).Checked = True
                        '    End If
                        'Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        ' aber nur, wenn es nicht vorher schon die leere Variante war 



                    End If

                    If aKtionskennung = PTTvActions.activateV Then
                        ' jetzt die Variante aktivieren 
                        Call replaceProjectVariant(pName, selectedVariantName, True, True, 0)
                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Call aktualisiereCharts(hproj, False)
                        Call awinNeuZeichnenDiagramme(2)
                    End If



            End Select

            stopRecursion = False

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            stopRecursion = True

            Select Case treeLevel

                Case 0 ' Projekt ist selektiert / nicht selektiert 

                    Dim selectedProjectName As String = node.Text
                    Dim selectedVariantName As String = ""

                    If node.Checked Then
                        ' wurde neu hinzugefügt 
                        ' war bereits vorher irgendwann mal eine Variante gewählt ?
                        Dim selectionExisted As Boolean = False
                        For j = 1 To node.Nodes.Count
                            childNode = node.Nodes.Item(j - 1)
                            If childNode.Checked Then
                                selectionExisted = True
                                selectedVariantName = getVariantNameOf(childNode.Text)
                            End If
                        Next

                        If Not selectionExisted And node.Nodes.Count > 0 Then
                            childNode = node.Nodes.Item(0)
                            childNode.Checked = True
                            selectedVariantName = getVariantNameOf(childNode.Text)
                        End If

                        ' jetzt muss das Projekt aus AlleProjekte auch in ShowProjekte transferiert werden 
                        Dim key As String = calcProjektKey(selectedProjectName, selectedVariantName)
                        Dim hproj As clsProjekt = AlleProjekte.getProject(key)

                        If Not ShowProjekte.contains(selectedProjectName) Then
                            ShowProjekte.Add(hproj)
                        End If

                        ' jetzt muss das Projekt neu gezeichnet werden ; 
                        ' dazu muss die Einfügestelle bestimmt werden, dann alle anderen Shapes nach unten verschoben werden 
                        ' hier muss die Zeile über Showprojekte bestimmt werden, einfach nach der Sortier-Reihenfolge 
                        ' das kann später dann noch angepasst werden 
                        Dim pZeile As Integer = ShowProjekte.getPTZeile(selectedProjectName)
                        'Dim pZeile2 As Integer = node.Index
                        'Call MsgBox("Zeile: " & pZeile.ToString)

                        If pZeile > 0 Then
                            Dim tmpCollection As New Collection
                            Call moveShapesDown(tmpCollection, pZeile, 1, 0)

                            Call ZeichneProjektinPlanTafel(tmpCollection, selectedProjectName, pZeile, tmpCollection, tmpCollection, True)
                        End If

                        ' jetzt muss noch geprüft werden , ob considerDependencies true ist 
                        If considerDependencies.Checked Then
                            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
                            Dim toDoListe As Collection = allDependencies.passiveListe(selectedProjectName, PTdpndncyType.inhalt)
                            If toDoListe.Count > 0 Then
                                For Each mprojectName As String In toDoListe
                                    Call activateMasterProject(mprojectName)
                                Next

                            End If
                        Else
                            ' nichts tun 
                        End If
                    Else
                        ' wurde abgewählt 
                        Dim pZeile As Integer

                        If ShowProjekte.contains(selectedProjectName) Then

                            pZeile = calcYCoordToZeile(projectboardShapes.getCoord(selectedProjectName)(0))
                            'pZeile = ShowProjekte.getPTZeile(selectedProjectName)
                            'Call MsgBox("Zeile: " & pZeile.ToString)

                            Call clearProjektinPlantafel(selectedProjectName)

                            ShowProjekte.Remove(selectedProjectName)


                            Call moveShapesUp(pZeile + 1, 1, True)

                        End If

                        ' jetzt muss noch geprüft werden , ob considerDependencies true ist 
                        If considerDependencies.Checked Then
                            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
                            Dim toDoListe As Collection = allDependencies.activeListe(selectedProjectName, PTdpndncyType.inhalt)
                            If toDoListe.Count > 0 Then
                                For Each dprojectName As String In toDoListe
                                    Call deactivateDependentProject(dprojectName)
                                Next

                            End If
                        Else
                            ' nichts tun 
                        End If
                    End If

                    ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
                    Call awinNeuZeichnenDiagramme(2)

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
                        ' bzw. was besser ist, den ersten Child-Knoten 
                        ' das funktioniert nämlich auch dann, wenn keine Variante mit Name "" existiert 
                        If projektNode.Nodes.Count > 0 Then
                            projektNode.Nodes.Item(0).Checked = True
                            selectedVariantName = getVariantNameOf(projektNode.Nodes.Item(0).Text)
                        Else
                            ' darf eigentlich gar nicht vorkommen 
                            selectedVariantName = ""
                        End If

                        ' '' die Standard Variante auf Checked setzen 
                        ''For i = 0 To projektNode.Nodes.Count - 1
                        ''    If projektNode.Nodes.Item(i).Text = "()" Then
                        ''        projektNode.Nodes.Item(i).Checked = True
                        ''    End If
                        ''Next

                        ' jetzt die selektierte Variante ins ShowProjekte stecken und aktualisieren ... 
                        ' aber nur, wenn es nicht vorher schon die leere Variante war 


                    End If

                    ' jetzt muss das bisherige aus ShowProjekte rausgenommen werden 
                    If ShowProjekte.contains(pName) And projektNode.Checked Then

                        Call replaceProjectVariant(pName, selectedVariantName, False, True, 0)
                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Call aktualisiereCharts(hproj, False)
                        Call awinNeuZeichnenDiagramme(2)



                        ' iwie Fehler 
                        ''ShowProjekte.Remove(pName)

                        ' '' jetzt muss das Projekt aus AlleProjekte auch in ShowProjekte transferiert werden 
                        ''Dim key As String = calcProjektKey(pName, selectedVariantName)
                        ''Dim hproj As clsProjekt = AlleProjekte.getProject(key)

                        ''If Not ShowProjekte.contains(pName) Then
                        ''    ShowProjekte.Add(hproj)
                        ''End If

                        ''Dim tmpCollection As New Collection
                        ''Call ZeichneProjektinPlanTafel(tmpCollection, pName, hproj.tfZeile, tmpCollection, tmpCollection, True)

                        ''Call awinNeuZeichnenDiagramme(2)
                    End If



            End Select

            stopRecursion = False

        End If


    End Sub

    ''' <summary>
    ''' aktiviert das Master-Projekt, wenn es nicht schon aktiviert ist ...
    ''' </summary>
    ''' <param name="mprojectName"></param>
    ''' <remarks></remarks>
    Private Sub activateMasterProject(ByVal mprojectName As String)

        For i As Integer = 1 To TreeViewProjekte.GetNodeCount(False)
            Dim curItem As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)

            If curItem.Checked Then
                ' nichts tun 
            Else
                stopRecursion = False
                curItem.Checked = True
                stopRecursion = True
            End If

        Next


    End Sub

    ''' <summary>
    ''' de-aktiviert das abhängige Projekt, wenn es nicht schon de-aktiviert ist 
    ''' </summary>
    ''' <param name="dprojectName"></param>
    ''' <remarks></remarks>
    Private Sub deactivateDependentProject(ByVal dprojectName As String)

        For i As Integer = 1 To TreeViewProjekte.GetNodeCount(False)
            Dim curItem As TreeNode = TreeViewProjekte.Nodes.Item(i - 1)

            If curItem.Text = dprojectName Then
                If curItem.Checked Then
                    stopRecursion = False
                    curItem.Checked = False
                    stopRecursion = True
                Else
                    ' nichts tun
                End If
            End If

        Next
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
                If aKtionskennung = PTTvActions.activateV Or _
                    (aKtionskennung = PTTvActions.chgInSession And node.Checked) Then
                    hproj = ShowProjekte.getProject(projName)
                ElseIf aKtionskennung = PTTvActions.chgInSession Then
                    ' jetzt erst noch die Variante bestimmen ... 
                    variantName = ""
                    For j As Integer = 1 To node.Nodes.Count
                        nodeVariant = node.Nodes.Item(j - 1)
                        If nodeVariant.Checked Then
                            variantName = nodeVariant.Text
                        End If
                    Next

                    Dim tmpKey As String = calcProjektKey(projName, variantName)
                    hproj = AlleProjekte.getProject(tmpKey)

                End If


                ' Löschen von Platzhalter
                node.Nodes.Clear()

                ' Eintragen der zum Projekt gehörenden Varianten
                For Each variantName In variantListe
                    nodeVariant = node.Nodes.Add(CType(variantName, String))

                    ' jetzt muss gecheckt werden , ob es sich um das Aktivieren handelt oder nicht
                    If aKtionskennung = PTTvActions.activateV Or _
                        aKtionskennung = PTTvActions.chgInSession Then
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
        Dim storedAtOrBefore As Date

        ' ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        ' ''Dim requestTrash As New Request(awinSettings.databaseURL, awinSettings.databaseName & "Trash", dbUsername, dbPasswort)

        If IsNothing(dropBoxTimeStamps.SelectedItem) Then
            storedAtOrBefore = Date.Now.AddDays(1)
        Else

            storedAtOrBefore = CDate(dropBoxTimeStamps.SelectedItem)

        End If

        Dim p As Integer, v As Integer, t As Integer

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
                                    Call loadProjectfromDB(pname, variantName, True, storedAtOrBefore)
                                Else
                                    Call loadProjectfromDB(pname, variantName, False, storedAtOrBefore)
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

                                    Call loadProjectfromDB(pname, variantName, first, storedAtOrBefore)
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

        ElseIf aKtionskennung = PTTvActions.chgInSession Then

            If dropboxScenarioNames.Text <> "" Then
                currentConstellation = dropboxScenarioNames.Text
                Call storeSessionConstellation(currentConstellation)
            End If

            DialogResult = Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("nicht unterstützte Option in ProjPortfolio Admin Formular ...")
        End If



    End Sub


    ''' <summary>
    ''' alle dargestellten Elemente im ProjektTree selektieren 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SelectionSet_Click(sender As Object, e As EventArgs) Handles SelectionSet.Click


        Dim projectNode As TreeNode

        stopRecursion = True

        With TreeViewProjekte

            ' die Behandlung von chgInSession ist etwas anders, weil sofort eine Aktion erfolgen muss ... 

            If aKtionskennung = PTTvActions.chgInSession Then

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    Dim pName As String = projectNode.Text

                    ' jetzt muss die Behandlung kommen, was denn gemacht werden soll 
                    ' ############ ChgInSession ####################################

                    ' das Projekt muss in Showprojekte, aber nur wenn es nicht bereits gecheckt war 
                    Dim variantName As String = ""

                    If Not projectNode.Checked Then
                        projectNode.Checked = True

                        ' ermittle die gecheckte Variante 
                        Dim checkedVariants As Collection = getNamesOfChildNodes(projectNode, True)
                        If checkedVariants.Count = 0 Then
                            ' dann muss der Varianten-Name entsprechend gesetzt werden  
                            Dim tmpCollection As Collection = AlleProjekte.getVariantNames(pName, True)
                            If tmpCollection.Count > 0 Then
                                variantName = getVariantNameOf(CStr(tmpCollection.Item(1)))
                            Else
                                variantName = ""
                            End If


                        ElseIf checkedVariants.Count = 1 Then
                            variantName = getVariantNameOf(CStr(checkedVariants.Item(1)))

                        ElseIf checkedVariants.Count > 1 Then
                            variantName = getVariantNameOf(CStr(checkedVariants.Item(1)))
                            For k As Integer = 1 To projectNode.Nodes.Count
                                If getVariantNameOf(projectNode.Nodes.Item(k - 1).Text) = variantName Then
                                    projectNode.Nodes.Item(k - 1).Checked = True
                                Else
                                    projectNode.Nodes.Item(k - 1).Checked = False
                                End If
                            Next
                        End If

                        ' jetzt muss das Projekt in Showprojekte eingetragen werden bzw. das alte zuvor gelöscht werden 
                        If ShowProjekte.contains(pName) Then
                            ShowProjekte.Remove(pName)
                        End If

                        Dim key As String = calcProjektKey(pName, variantName)
                        Dim hproj As clsProjekt = AlleProjekte.getProject(key)

                        ShowProjekte.Add(hproj)

                    Else
                        ' nichts tun , denn das Projekt wird bereits angezeigt und ist in Showprojekte drin 
                    End If

                Next

                ' jetzt im formular den Mauszeiger auf Warten ... setzen 
                Me.Cursor = Cursors.WaitCursor

                ' jetzt muss die Plan-Tafel gelöscht werden 
                Call awinClearPlanTafel()

                ' jetzt muss die Plan-Tafel neu gezeichnet werden 
                Call awinZeichnePlanTafelNeu(True)

                ' jetzt müssen die Diagramme neu gezeichnet werden 
                Call awinNeuZeichnenDiagramme(2)

                Me.Cursor = Cursors.Default

            ElseIf aKtionskennung = PTTvActions.deleteV Or _
                aKtionskennung = PTTvActions.activateV Then
                ' nichts tun, Alle Selektieren macht bei diesen keinen Sinn 

            ElseIf aKtionskennung = PTTvActions.delFromDB Then
                ' wenn ein Stand angegeben ist , dann sollen alle mit diesem Stand markiert werden 
                If Not IsNothing(dropBoxTimeStamps.SelectedItem) Then

                    Dim lookForTimestamp As Date = CDate(dropBoxTimeStamps.SelectedItem)
                    Dim vergleichsString As String = lookForTimestamp.ToString

                    ' jetzt wird der TreeView komplett expanded ...
                    stopRecursion = False
                    .ExpandAll()
                    stopRecursion = True

                    For i As Integer = 1 To .Nodes.Count
                        projectNode = .Nodes.Item(i - 1)

                        For v As Integer = 1 To projectNode.Nodes.Count
                            Dim variantNode As TreeNode = projectNode.Nodes.Item(v - 1)

                            For t As Integer = 1 To variantNode.Nodes.Count
                                Dim tsNode As TreeNode = variantNode.Nodes.Item(t - 1)
                                If tsNode.Text = vergleichsString Then
                                    tsNode.Checked = True
                                    variantNode.Checked = False
                                    projectNode.Checked = False

                                    If Not projectNode.IsExpanded Then
                                        projectNode.Expand()
                                    End If

                                    If Not variantNode.IsExpanded Then
                                        variantNode.Expand()
                                    End If
                                End If
                            Next
                        Next

                    Next
                Else
                    Call MsgBox("nur aktiv in Verbindung mit einem ausgewählten Stand")
                End If
                
            Else
                ' in allen anderen Fällen: loadPV, loadPVS, delFromDB, delFromSession

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    If Not projectNode.Checked Then
                        projectNode.Checked = True
                    End If
                    If projectNode.Nodes.Count > 0 Then
                        Call Check(projectNode)
                    End If
                Next

            End If

        End With

        stopRecursion = False

    End Sub

    ''' <summary>
    ''' gibt die Namen der Kind-Knoten wieder, 
    ''' alle
    ''' nur die, die gecheckt sind  
    ''' nur die, die nicht gecheckt sind 
    ''' </summary>
    ''' <param name="curNode">der aktuelle Knoten </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getNamesOfChildNodes(ByVal curNode As TreeNode, ByVal checkState As Boolean, Optional considerAll As Boolean = False) As Collection
        Dim tmpCollection As New Collection

        Dim childNode As TreeNode

        With curNode

            For i As Integer = 1 To .Nodes.Count

                childNode = .Nodes.Item(i - 1)

                If considerAll Then
                    tmpCollection.Add(childNode.Name)
                Else
                    If childNode.Checked = checkState Then
                        tmpCollection.Add(childNode.Name)
                    End If
                End If
            Next

        End With

        getNamesOfChildNodes = tmpCollection
    End Function

    ''' <summary>
    ''' setzt alle Knoten im TreeView auf checked
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub Check(ByRef node As TreeNode)
        Dim curNode As TreeNode

        With node

            For i As Integer = 1 To .Nodes.Count
                curNode = .Nodes.Item(i - 1)
                If Not curNode.Checked Then
                    curNode.Checked = True
                End If
                If curNode.Nodes.Count > 0 Then
                    Call Check(curNode)
                End If
            Next

        End With

    End Sub

    ''' <summary>
    ''' setzt alle Knoten im TreeView auf unchecked
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub unCheck(ByRef node As TreeNode)
        Dim curNode As TreeNode

        With node

            For i As Integer = 1 To .Nodes.Count
                curNode = .Nodes.Item(i - 1)
                If curNode.Checked Then
                    curNode.Checked = False
                End If
                If curNode.Nodes.Count > 0 Then
                    Call unCheck(curNode)
                End If
            Next

        End With

    End Sub


    Private Sub expandCompletely_Click(sender As Object, e As EventArgs) Handles expandCompletely.Click

        With TreeViewProjekte
            .ExpandAll()
        End With

    End Sub

    Private Sub collapseCompletely_Click(sender As Object, e As EventArgs) Handles collapseCompletely.Click
        With TreeViewProjekte
            .CollapseAll()
        End With
    End Sub

    ''' <summary>
    ''' de-selektiert alle Knoten im Formular ProjektPortfolio 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SelectionReset_Click(sender As Object, e As EventArgs) Handles SelectionReset.Click

        Dim projectNode As TreeNode

        stopRecursion = True

        With TreeViewProjekte

            ' die Behandlung von chgInSession ist etwas anders, weil sofort eine Aktion erfolgen muss ... 

            If aKtionskennung = PTTvActions.chgInSession Then

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    Dim pName As String = projectNode.Text

                    ' jetzt muss die Behandlung kommen, was denn gemacht werden soll 
                    ' ############ ChgInSession ####################################

                    ' das Projekt muss in Showprojekte, aber nur wenn es nicht bereits gecheckt war 
                    Dim variantName As String = ""

                    If projectNode.Checked Then

                        projectNode.Checked = False

                    Else
                        ' nichts tun , denn das Projekt wird bereits angezeigt und ist in Showprojekte drin 
                    End If

                Next

                ' jetzt muss die Plan-Tafel gelöscht werden 
                Call awinClearPlanTafel()

                ' jetzt muss Showprojekte gelöscht werden 
                ShowProjekte.Clear()

                ' jetzt müssen die Diagramme neu gezeichnet werden 
                Call awinNeuZeichnenDiagramme(2)

            ElseIf aKtionskennung = PTTvActions.activateV Then
                ' nichts tun, Alle Resetten macht bei diesen keinen Sinn 

            Else
                ' auch in den Fällen deleteV
                ' in allen anderen Fällen: loadPV, loadPVS, delFromDB, delFromSession

                For i As Integer = 1 To .Nodes.Count
                    projectNode = .Nodes.Item(i - 1)
                    If projectNode.Checked Then
                        projectNode.Checked = False
                    End If

                    If projectNode.Nodes.Count > 0 Then
                        Call unCheck(projectNode)
                    End If
                Next

            End If

        End With

        stopRecursion = False


    End Sub

    
    
    
    ''' <summary>
    ''' reduziert anhand des definierten Filters die aktuelle Gesamtliste und baut den Treeview wieder auf 
    ''' 
    '''
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub filterIcon_Click(sender As Object, e As EventArgs) Handles filterIcon.Click

        Dim filterFormular As New frmNameSelection

        With filterFormular
            .menuOption = PTmenue.filterdefinieren
            If .ShowDialog() = Windows.Forms.DialogResult.OK Then

                stopRecursion = True
                Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, aKtionskennung, _
                                   True, Date.Now)
                stopRecursion = False

                ' Das DeleteFilterIcon mit Bild versehen 
                Me.deleteFilterIcon.Image = My.Resources.funnel_delete
                Me.deleteFilterIcon.Enabled = True

            End If
        End With
    End Sub


    Private Sub dropBoxTimeStamps_MouseHover(sender As Object, e As EventArgs) Handles dropBoxTimeStamps.MouseHover
        ToolTipStand.Show("Auswahl eines Zeitstempels; im Default wird immer der letzte Stand berücksichtigt", dropBoxTimeStamps, 2000)
    End Sub

    Private Sub dropBoxTimeStamps_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropBoxTimeStamps.SelectedIndexChanged

    End Sub

    Private Sub dropboxScenarioNames_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dropboxScenarioNames.SelectedIndexChanged

    End Sub

    Private Sub deleteFilterIcon_Click(sender As Object, e As EventArgs) Handles deleteFilterIcon.Click

        aktuelleGesamtListe = AlleProjekte

        stopRecursion = True
        Call buildTreeview(projektHistorien, TreeViewProjekte, aktuelleGesamtListe, aKtionskennung, _
                           False, Date.Now)
        stopRecursion = False

            ' Das DeleteFilterIcon mit Bild versehen 
        Me.deleteFilterIcon.Image = Nothing
        Me.deleteFilterIcon.Enabled = False

        
    End Sub
End Class