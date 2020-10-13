Imports DBAccLayer
Imports ProjectBoardDefinitions
Imports System.Windows.Forms
Imports System.Drawing
Public Class frmLoadConstellation

    Private formerselect As String
    Private stopRecursion As Boolean = False
    Public retrieveFromDB As Boolean
    Public earliestDate As Date
    Private actionID As Integer
    Public constellationsToShow As SortedList(Of String, String)

    Private Sub frmLoadConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Call languageSettings()
        If constellationsToShow.Count > 0 Then

            For Each kvp As KeyValuePair(Of String, String) In constellationsToShow
                With TreeViewPortfolios
                    .CheckBoxes = True
                    .Nodes.Add(kvp.Key)
                End With

            Next

            stopRecursion = False
            If Not retrieveFromDB Then
                requiredDate.Visible = False
                lblStandvom.Visible = False
                updateTreeview(constellationsToShow, actionID, False)
            Else
                updateTreeview(constellationsToShow, actionID, True)
            End If
        Else
            DialogResult = System.Windows.Forms.DialogResult.Cancel
            MyBase.Close()
        End If

        formerselect = ""

    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            lblStandvom.Text = "Version"
            addToSession.Text = "add to session"
            OKButton.Text = "OK"
            Abbrechen.Text = "Cancel"
            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                If awinSettings.loadPFV Then
                    loadAsSummary.Text = "load and show summary project"
                Else
                    loadAsSummary.Text = "calculate and show summary project"
                End If
            Else
                loadAsSummary.Text = "calculate and show summary project"
            End If
        Else
            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                If awinSettings.loadPFV Then
                    loadAsSummary.Text = "Summary Projekt laden und anzeigen"
                Else
                    loadAsSummary.Text = "Summary Projekt berechnen und anzeigen"
                End If
            Else
                loadAsSummary.Text = "Summary Projekt berechnen und anzeigen"
            End If
        End If

    End Sub

    Private Sub Abbrechen_Click(sender As Object, e As EventArgs) Handles Abbrechen.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()
    End Sub



    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click


        If TreeViewPortfolios.Nodes.Count >= 1 Then
            DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("bitte einen Eintrag selektieren")
        End If

    End Sub

    Private Sub addToSession_CheckedChanged(sender As Object, e As EventArgs) Handles addToSession.CheckedChanged


    End Sub

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        retrieveFromDB = False
        'constellationsToShow = New clsConstellations
        constellationsToShow = New SortedList(Of String, String)

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Private Sub dropBoxTimeStamps_SelectedIndexChanged(sender As Object, e As EventArgs)

        ' den Fokus von diesem Element wegnehmen 
        TreeViewPortfolios.Focus()
        Try
            With TreeViewPortfolios
                .CollapseAll()
            End With
        Catch ex As Exception

        End Try

    End Sub

    Private Sub requiredDate_ValueChanged(sender As Object, e As EventArgs) Handles requiredDate.ValueChanged

        If Not IsNothing(requiredDate) Then

            If requiredDate.Value >= earliestDate Then
                requiredDate.Value = requiredDate.Value.Date.AddHours(23).AddMinutes(59)
            Else
                Call MsgBox("es gibt vor dem " & earliestDate.ToShortDateString & " keine Projekte in der Datenbank ")
                requiredDate.Value = Date.Now.Date.AddHours(23).AddMinutes(59)
            End If

        Else
            ' nichts tun ...
        End If

    End Sub

    Private Sub TreeViewPortfolios_AfterSelect(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeViewPortfolios.AfterSelect

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


    ''' <summary>
    ''' aktualisiert bzw. baut die TreeView gemäß der aktuelleGesamtListe bzw. der pvNamesList neu auf
    ''' Rahmenbedingung: stopRecursion ist immer False, wenn Update TreeView aufgerufen wird 
    ''' </summary>
    ''' <param name="pvNamesList"></param>
    ''' <param name="aKtionskennung"></param>
    ''' <param name="quickList"></param>
    ''' <remarks></remarks>
    Private Sub updateTreeview(ByVal pvNamesList As SortedList(Of String, String),
                                  ByVal aKtionskennung As Integer,
                                  ByVal quickList As Boolean)

        Dim err As New clsErrorCodeMsg

        Dim portfolioNode As TreeNode
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
        'Dim storedHeute As Date = Now
        Dim storedGestern As Date = StartofCalendar
        Dim pname As String = ""
        Dim vpid As String = ""
        Dim variantName As String = ""
        Dim loadErrorMsg As String = ""

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' es ist ein Zeitraum definiert 
            zeitraumVon = getDateofColumn(showRangeLeft, False)
            zeitraumbis = getDateofColumn(showRangeRight, True)
        End If

        ' steuert, ob erstmal nur Projekt-Namen, Varianten-Namen gelesen werden 
        ' geht wesentlich schneller, wenn es sich um eine Datenbank mit sehr vielen Projekten handelt ... 


        With TreeViewPortfolios
            .Nodes.Clear()
        End With

        If pvNamesList.Count >= 1 Then

            'If Not noDB And aKtionskennung = PTTvActions.setWriteProtection Then
            '    'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            '    writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
            'End If

            With TreeViewPortfolios

                .CheckBoxes = True

                Dim portfolioliste As SortedList(Of String, String)

                If quickList Then
                    portfolioliste = New SortedList(Of String, String)
                    For Each kvp As KeyValuePair(Of String, String) In pvNamesList
                        Dim tmpName As String = kvp.Key
                        Dim tmpVpid As String = kvp.Value
                        If Not portfolioliste.ContainsKey(tmpName) Then
                            portfolioliste.Add(tmpName, tmpVpid)
                        End If
                    Next

                Else
                    ' hole die Namen aus der sortierten Liste, nicht aus der cItem-Liste 
                    portfolioliste = constellationsToShow
                End If

                Dim showPname As Boolean



                For Each kvp As KeyValuePair(Of String, String) In portfolioliste

                    showPname = True

                    pname = getPnameFromKey(kvp.Key) ' der key ist das sortier-Kriterium, kann constellationName sein, aber auch was ganz anderes 


                    Dim hPortfolio As clsConstellation = Nothing
                    Dim variantNames As New Collection

                    If quickList Then
                        vpid = kvp.Value
                        variantNames = getVariantListeFromPName(pname, vpid, ptPRPFType.portfolio)
                        variantNames.Add("")    ' Standard-Variante hinzufügen
                    Else
                        variantName = getVariantnameFromKey(kvp.Key)
                        variantNames.Add(variantName)
                    End If


                    If showPname Then


                        portfolioNode = .Nodes.Add(pname)


                        ' damit kann evtl direkt auf den Node zugegriffen werden ...
                        portfolioNode.Name = pname



                        If Not IsNothing(hPortfolio) Then
                            variantName = hPortfolio.variantName
                        End If

                        ' Platzhalter einfügen; wird für alle Aktionskennungen benötigt

                        If variantNames.Count > 0 Then

                            Dim vName As String = variantName
                            portfolioNode.Tag = "X"
                            For iv As Integer = 1 To variantNames.Count
                                vName = CStr(variantNames.Item(iv))
                                Dim vNameStripped As String = ""
                                Dim tmpStr() As String = vName.Split(New Char() {CChar("("), CChar(")")})
                                If tmpStr.Length = 1 Then
                                    vNameStripped = tmpStr(0)
                                ElseIf tmpStr.Length >= 3 Then
                                    vNameStripped = tmpStr(1).Trim
                                End If
                                ' pfv-Variante wird nicht in den Tree mit aufgenommen
                                If vName <> ptVariantFixNames.pfv.ToString Then
                                    Dim variantNode As TreeNode = portfolioNode.Nodes.Add(vName)
                                    variantNode.Text = "(" & vName & ")"
                                    variantNode.Tag = "X"
                                End If

                                'If aKtionskennung = PTTvActions.delFromDB Then
                                '    variantNode.Tag = "P"
                                '    Dim tmpNodeLevel2 As TreeNode = variantNode.Nodes.Add("Platzhalter-Datum")
                                'Else
                                '    variantNode.Tag = "X"
                                'End If

                                'Call bestimmeNodeCheckStatus(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant,
                                '                             pname, vNameStripped)
                                'Call bestimmeNodeAppearance(variantNode, aKtionskennung, PTTreeNodeTyp.pVariant, pname, vNameStripped)

                            Next

                        Else
                            portfolioNode.Tag = "X"
                        End If

                        'Call bestimmeNodeCheckStatus(portfolioNode, aKtionskennung, PTTreeNodeTyp.project,
                        '                              pname, variantName)
                        'Call bestimmeNodeAppearance(portfolioNode, aKtionskennung, PTTreeNodeTyp.project, pname, variantName)

                    End If

                Next

            End With
        Else
            If awinSettings.englishLanguage Then
                loadErrorMsg = "No Portfolios loaded!"
            Else
                loadErrorMsg = "Es sind keine Portfolios in der Session geladen!"
            End If
            Call MsgBox(loadErrorMsg)
        End If


    End Sub


    Private Sub TreeViewPortfolios_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewPortfolios.AfterCheck

        Dim node As TreeNode
        Dim schluessel As String = ""
        Dim treeLevel As Integer
        Dim currentIndex As Integer
        Dim lastlevelChecked As Integer
        Dim lastIndexChecked As Integer
        Dim shiftKeywasPressed As Boolean = False


        ' Andernfalls wird die Check Routine endlos aufgerufen ...
        If stopRecursion Then
            Exit Sub
        End If

        node = e.Node
        treeLevel = node.Level
        currentIndex = node.Index

        ' hier wird jetzt sichergestellt, daß nur die nach der aktuellen Aktion gültigen Checks gesetzt werden können
        ' vor allem muss überall dort, wo das Szenario mit diesem Check verändert wird, das currentBrowserSzenario geupdated werden ...
        ' mit Click in TreeView wird verändert: Activate Variant, ChgInSession 

        Dim checkMode As Boolean = node.Checked

        stopRecursion = True

        Select Case treeLevel

            Case 0 ' Portfolio ist selektiert / nicht selektiert 

                ' es  dürfen mehrere Portfolios gecheckt sein

                For h = 0 To TreeViewPortfolios.Nodes.Count - 1
                    Dim tmpNode As TreeNode = TreeViewPortfolios.Nodes.Item(h)

                    If (tmpNode.Level = treeLevel) And (tmpNode.Text = node.Text) Then
                        ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                        tmpNode.Checked = checkMode
                        'Call collectAfterCheck(treeLevel, tmpNode)
                        If Not tmpNode.Checked Then
                            For Each vNode As TreeNode In tmpNode.Nodes
                                vNode.Checked = False
                            Next
                        End If
                    End If
                Next

            Case 1 ' Variante ist selektiert / nicht selektiert

                ' es darf nur eine Variante gecheckt sein

                For h = 0 To node.Parent.Nodes.Count - 1
                    Dim tmpNode As TreeNode = node.Parent.Nodes.Item(h)
                    If (tmpNode.Level = treeLevel) And (tmpNode.Text = node.Text) Then
                        ' Aktion nur durchführen, wenn auf der gleichen Ebene 
                        tmpNode.Checked = checkMode
                        If tmpNode.Checked Then
                            node.Parent.Checked = True
                        End If
                        'Call collectAfterCheck(treeLevel, tmpNode)
                    Else
                        ' uncheck alle anderen Varianten
                        tmpNode.Checked = False
                    End If


                Next

        End Select

        stopRecursion = False

        ' merken , wo zum letzten Mal geklickt wurde ....
        lastLevelChecked = treeLevel
        lastIndexChecked = currentIndex

    End Sub

    ''' <summary>
    ''' führt die Aktion aus .. wird jetzt benötigt, um mit Shift mehrere Aktionen gleichzeitig durchführen zu können 
    ''' </summary>
    ''' <param name="TreeLevel"></param>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Private Sub collectAfterCheck(ByVal TreeLevel As Integer, ByVal node As TreeNode)

        Dim err As New clsErrorCodeMsg

        Dim childNode As TreeNode
        Dim parentNode As TreeNode

        Select Case TreeLevel

            Case 0 ' Projekt ist selektiert / nicht selektiert 

                Dim checkMode As Boolean = node.Checked

                For i = 1 To node.Nodes.Count
                    ' Schleife über alle Varianten
                    childNode = node.Nodes.Item(i - 1)
                    childNode.Checked = checkMode

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

                ' wenn mit diesem Knoten jetzt alle geckecked/unchecked sind, soll auch parent wieder gesetzt werden 
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

    End Sub


    Private Sub TreeViewPortfolios_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TreeViewPortfolios.BeforeExpand


        Dim err As New clsErrorCodeMsg

        Dim selectedNode As New TreeNode
        Dim variantNode As New TreeNode
        Dim nodeTimeStamp As New TreeNode
        Dim pName As String = ""
        Dim variantName As String = ""
        'Dim hliste As SortedList(Of Date, String)
        Dim nodeLevel As Integer
        Dim variantListe As Collection
        Dim hportfolio As New clsConstellation

        selectedNode = e.Node
        nodeLevel = e.Node.Level

        ' Projekt-Ebene
        If nodeLevel = 0 Then


            pName = selectedNode.Text

            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If selectedNode.Tag = "P" Then

                'Call MsgBox("sollte eigentlich gar nicht mehr vorkommen ...")
                ' Inhalte der Sub-Nodes müssen neu aufgebaut werden 

                variantListe = getVariantListeFromPName(pName, "", ptPRPFType.portfolio)


                ' Löschen von Platzhalter
                selectedNode.Nodes.Clear()

                ' wird nur im Fall loadPV benötigt ... wird baer hier besetzt, weil das sonst in der Schleife ständig neu ausgerechnet wird 
                ' es wird entweder by default die Basis Variante geladen, oder die pfv-Variante, sofern das so gesetzt ist und die customUserRole = PMGR
                Dim variantNameLookedFor As String = ""
                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV = True Then
                    variantNameLookedFor = ptVariantFixNames.pfv.ToString
                Else
                    variantNameLookedFor = ""
                End If

                ' Eintragen der zum Portfolio gehörenden Varianten
                For Each variantName In variantListe
                    variantNode = selectedNode.Nodes.Add(CType(variantName, String))

                    ' es wird by default nur eine Projekt-Variante selektiert ...

                    stopRecursion = True

                    ' es wird nur gecheckt, wenn selectedNode.checked 
                    If selectedNode.Checked = True Then
                        ' entscheiden, welche Variante gecheckt wird 
                        If variantName = variantNameLookedFor Then
                            variantNode.Checked = True
                        End If
                    Else
                        ' auf jeden Fall nicht checken ..
                        variantNode.Checked = False
                    End If

                    stopRecursion = False
                Next

                selectedNode.Tag = "X"
            End If



        ElseIf nodeLevel = 1 Then

            ' hier wurde eine Variante selektiert ...

            If selectedNode.Tag = "P" Then

                selectedNode.Tag = "X"
                pName = selectedNode.Parent.Text
                variantName = selectedNode.Text

            End If


        End If





    End Sub
End Class