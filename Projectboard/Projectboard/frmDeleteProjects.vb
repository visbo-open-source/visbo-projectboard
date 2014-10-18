Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports System.Windows.Forms



Public Class frmDeleteProjects

    ' Public projekteInDB As New SortedList(Of String, clsProjekt)
    Private aktuelleGesamtListe As New clsProjekteAlle
    Private projektHistorien As New clsProjektDBInfos
    Private stopRecursion As Boolean = False
    ' wird an der aufrufenden Stelle gesetzt; steuert, was mit den ausgewählten ELementen geschieht
    Friend aKtionskennung As Integer
    Friend selectedItems As New clsProjektDBInfos

    
    Private Sub frmDeleteProjects_FormClosed(sender As Object, e As EventArgs) Handles Me.FormClosed

        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left
        projektHistorien.clear()

    End Sub



    Private Sub frmDeleteProjects_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Call buildTreeview()

        'Dim nodeLevel0 As TreeNode
        'Dim nodeLevel1 As TreeNode
        'Dim zeitraumVon As Date = StartofCalendar
        'Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
        'Dim storedHeute As Date = Now
        'Dim storedGestern As Date = StartofCalendar
        'Dim pname As String = ""
        'Dim variantName As String = ""
        'Dim loadErrorMsg As String = ""


        'Dim deletedProj As Integer = 0
        ''Dim singleShp As Excel.Shape
        ''Dim awinSelection As Excel.ShapeRange
        ''Dim anzElements As Integer
        ''Dim hproj As clsProjekt
        'Dim schluessel As String = ""

        'Dim request As New Request(awinSettings.databaseName)
        'Dim requestTrash As New Request(awinSettings.databaseName & "Trash")

        'projektHistorien.clear()

        '' Alle Projekte aus DB
        '' projekteInDB = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)

        'Select Case aKtionskennung

        '    Case PTtvactions.delFromDB
        '        pname = ""
        '        variantName = ""
        '        aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
        '        loadErrorMsg = "es gibt keine Projekte in der Datenbank"

        '    Case PTtvactions.delFromSession
        '        aktuelleGesamtListe = AlleProjekte
        '        loadErrorMsg = "es sind keine Projekte geladen"

        '    Case PTtvactions.loadPVS
        '        pname = ""
        '        variantName = ""
        '        aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
        '        loadErrorMsg = "es gibt keine Projekte in der Datenbank"

        '    Case PTtvactions.activateV
        '        aktuelleGesamtListe = AlleProjekte
        '        loadErrorMsg = "es sind keine Projekte geladen"

        'End Select


        'If aktuelleGesamtListe.Count > 1 Then

        '    With TreeViewProjekte

        '        .CheckBoxes = True

        '        Dim projektliste As Collection = aktuelleGesamtListe.getProjectNames

        '        For Each pname In projektliste

        '            nodeLevel0 = .Nodes.Add(pname)
        '            nodeLevel0.Tag = "P"
        '            Dim variantListe As Collection = aktuelleGesamtListe.getVariantNames(pname)

        '            ' Platzhalter einfügen
        '            nodeLevel1 = nodeLevel0.Nodes.Add("()")
        '            nodeLevel1.Tag = "P"

        '        Next


        '    End With
        'Else
        '    Call MsgBox(loadErrorMsg)
        'End If


        ' Beginn alter Code 
        '
        '
        'Try
        '    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        'Catch ex As Exception
        '    awinSelection = Nothing
        'End Try

        'If Not awinSelection Is Nothing Then    ' es sind Projekte selektiert

        '    If awinSelection.Count > 0 Then
        '        'selektierte Projekte ins Formular eintragen
        '        anzElements = awinSelection.Count
        '        Dim i As Integer
        '        For i = 1 To anzElements

        '            singleShp = awinSelection.Item(i)
        '            hproj = ShowProjekte.getProject(singleShp.Name)


        '            schluessel = calcProjektKey(hproj)


        '            projektHistorien.Add(schluessel, Date.MinValue) 'Platzhalter für die Projekthistorie

        '            With TreeViewProjekte

        '                .CheckBoxes = True

        '                nodeLevel0 = .Nodes.Add(hproj.name)
        '                nodeLevel1 = nodeLevel0.Nodes.Add(hproj.variantName)
        '                nodeLevel1.Nodes.Add(CType(Date.MinValue, String))    'Platzhalter für die Projekthistorie

        '            End With
        '        Next i

        '    End If



        'Else


        '    '' geladene Projekte ins Formular eintragen

        '    If AlleProjekte.Count > 0 Then

        '        With TreeViewProjekte

        '            .CheckBoxes = True

        '            Dim projektliste As Collection = AlleProjekte.getProjectNames

        '            For Each pname In projektliste

        '                nodeLevel0 = .Nodes.Add(pname)
        '                With nodeLevel0
        '                    .ToolTipText = "Projekt-Name"
        '                End With

        '                Dim variantListe As Collection = AlleProjekte.getVariantNames(pname)

        '                For Each vName As String In variantListe
        '                    nodeLevel1 = nodeLevel0.Nodes.Add(vName)
        '                    With nodeLevel0
        '                        .ToolTipText = "Varianten-Name"
        '                    End With
        '                    nodeLevel1.Nodes.Add(CType(Date.MinValue, String))    'Platzhalter für die Projekthistorie
        '                Next

        '            Next

        '        End With



        '    Else
        '        ' Alle Projekte aus DB



        '        If aktuelleGesamtListe.Count > 1 Then

        '            With TreeViewProjekte

        '                .CheckBoxes = True

        '                Dim projektliste As Collection = aktuelleGesamtListe.getProjectNames

        '                For Each pname In projektliste

        '                    nodeLevel0 = .Nodes.Add(pname)
        '                    nodeLevel0.Tag = "P"
        '                    Dim variantListe As Collection = aktuelleGesamtListe.getVariantNames(pname)

        '                    ' Platzhalter einfügen
        '                    nodeLevel1 = nodeLevel0.Nodes.Add("()")
        '                    nodeLevel1.Tag = "P"

        '                Next


        '            End With
        '        Else
        '            Call MsgBox(" keine Projekte in der Datenbank")
        '        End If

        '    End If 'AlleProjekte


        'End If 'selektierte Projekte
        ' Ende alter Code

    End Sub

    
    ''' <summary>
    ''' Aktion, die ausgeführt wird, nachdem eine Checkbox gewählt oder abgewählt wurde 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TreeViewProjekte_AfterCheck(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeViewProjekte.AfterCheck
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

        Select Case aKtionskennung

            Case PTtvactions.delFromDB

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

            Case PTtvactions.delFromSession

            Case PTtvactions.loadPVS

            Case PTtvactions.activateV

        End Select


        ' alter Code
        'schluessel = CType(node.Text, String) & "#"

        'If Not IsNothing(node.Parent) Then
        '    schluessel = CType(node.Parent.Text, String) & "#"
        '    Try
        '        selCollection = projektHistorien.getTimeStamps(schluessel)

        '        ' Löschen aus der projektHistorien-Liste
        '        ' projektHistorien.Remove(schluessel, CType(node.Text, Date))
        '        If node.Checked = True Then
        '            ' Aufbau der Liste selectedToDelete
        '            selectedItems.Add(schluessel, selCollection.ElementAt(node.Index).Key)
        '        Else
        '            selectedItems.Remove(schluessel, selCollection.ElementAt(node.Index).Key)
        '        End If
        '    Catch ex As Exception

        '    End Try


        'Else
        '    schluessel = CType(node.Text, String) & "#"
        '    Try


        '        If node.Checked = True Then

        '            If node.IsExpanded = False Then
        '                node.Expand()
        '            End If

        '            selCollection = projektHistorien.getTimeStamps(schluessel)

        '            Dim i As Integer
        '            For i = 1 To selCollection.Count
        '                timeStamp = selCollection.ElementAt(i - 1).Key
        '                selectedItems.Add(schluessel, timeStamp)

        '                'Alle Unterknoten werden zum Löschen gecheckt
        '                e.Node.Nodes(i - 1).Checked = True

        '            Next i
        '        Else

        '            selCollection = projektHistorien.getTimeStamps(schluessel)

        '            Dim i As Integer
        '            For i = 1 To selCollection.Count
        '                timeStamp = selCollection.ElementAt(i - 1).Key
        '                selectedItems.Remove(schluessel, timeStamp)

        '                '' Check wird für alle Unterknoten entfernt
        '                e.Node.Nodes(i - 1).Checked = False

        '            Next i
        '        End If

        '    Catch ex As Exception

        '    End Try


        'End If

    End Sub
    Private Sub TreeViewProjekte_AfterCollapse(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeViewProjekte.AfterCollapse

    End Sub
    Private Sub TreeViewProjekte_BeforeExpand(sender As Object, e As Windows.Forms.TreeViewCancelEventArgs) Handles TreeViewProjekte.BeforeExpand

        Dim request As New Request(awinSettings.databaseName)
        Dim node As New TreeNode
        Dim nodeVariant As New TreeNode
        Dim nodeTimeStamp As New TreeNode
        Dim projName As String = ""
        Dim variantName As String = ""
        Dim hliste As SortedList(Of Date, String)
        Dim nodeLevel As Integer
        Dim variantListe As Collection



        node = e.Node
        nodeLevel = node.Level

        If nodeLevel = 0 Then
            projName = node.Text

            ' node.tag = P bedeutet, daß es sich noch um einen Platzhalter handelt 
            If node.Tag = "P" Then
                ' Inhalte der Sub-Nodes müssen neu aufgebaut werden 
                variantListe = aktuelleGesamtListe.getVariantNames(projName)

                ' Löschen von Platzhalter
                node.Nodes.Clear()

                ' Eintragen der zum Projekt gehörenden Varianten
                For Each variantName In variantListe
                    nodeVariant = node.Nodes.Add(CType(variantName, String))
                    nodeVariant.Checked = node.Checked


                    If aKtionskennung = 0 Or _
                        aKtionskennung = 2 Then
                        ' Einfügen eines Platzhalters macht nur Sinn bei Snapshots löschen bzw. Snapshots laden 

                        nodeVariant.Tag = "P"
                        nodeVariant.Nodes.Add("()")
                    Else
                        nodeVariant.Tag = "X"
                    End If
                    
                Next

                node.Tag = "X"
                'If node.IsSelected Then
                '    node.Expand()
                'End If
            End If



        ElseIf nodeLevel = 1 Then


            If node.Tag = "P" Then

                node.Tag = "X"
                projName = node.Parent.Text
                variantName = getVariantNameOf(node.Text)

                hliste = projektHistorien.getTimeStamps(calcProjektKey(projName, variantName))

                If hliste.Count = 0 Then

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

                        'If node.IsSelected Then
                        '    node.Expand()
                        'End If
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
        If tmpstr.Count = 3 Then
            vName = tmpstr(1)
        End If

        getVariantNAmeOf = vName

    End Function

    ''' <summary>
    ''' baut den aktuell gültigen Treeview auf  
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub buildTreeview()

        Dim nodeLevel0 As TreeNode
        Dim nodeLevel1 As TreeNode
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = StartofCalendar
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim loadErrorMsg As String = ""


        Dim deletedProj As Integer = 0


        Dim request As New Request(awinSettings.databaseName)
        Dim requestTrash As New Request(awinSettings.databaseName & "Trash")

        ' alles zurücksetzen 
        projektHistorien.clear()

        With TreeViewProjekte
            .Nodes.Clear()
        End With

        ' Alle Projekte aus DB
        ' projekteInDB = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)

        Select Case aKtionskennung

            Case PTtvactions.delFromDB
                pname = ""
                variantName = ""
                aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
                loadErrorMsg = "es gibt keine Projekte in der Datenbank"

            Case PTtvactions.delFromSession
                aktuelleGesamtListe = AlleProjekte
                loadErrorMsg = "es sind keine Projekte geladen"

            Case PTtvactions.loadPVS
                pname = ""
                variantName = ""
                aktuelleGesamtListe.liste = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
                loadErrorMsg = "es gibt keine Projekte in der Datenbank"

            Case PTtvactions.activateV
                aktuelleGesamtListe = AlleProjekte
                loadErrorMsg = "es sind keine Projekte geladen"

        End Select


        If aktuelleGesamtListe.Count > 1 Then

            With TreeViewProjekte

                .CheckBoxes = True

                Dim projektliste As Collection = aktuelleGesamtListe.getProjectNames

                For Each pname In projektliste

                    nodeLevel0 = .Nodes.Add(pname)
                    nodeLevel0.Tag = "P"
                    Dim variantListe As Collection = aktuelleGesamtListe.getVariantNames(pname)

                    ' Platzhalter einfügen
                    nodeLevel1 = nodeLevel0.Nodes.Add("()")
                    nodeLevel1.Tag = "P"

                Next


            End With
        Else
            Call MsgBox(loadErrorMsg)
        End If


    End Sub


    ''' <summary>
    ''' wird bei Auslösen des "Aktionsbuttons" ausgeführt; 
    ''' in Abhängigkeit von Aktionskennung 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SubmitButton_Click(sender As Object, e As EventArgs) Handles SubmitButton.Click

        Dim projektNode As TreeNode, variantNode As TreeNode, timeStampNode As TreeNode
        Dim anzahlProjekte As Integer
        Dim anzahlVarianten As Integer
        Dim anzahlTimeStamps As Integer
        Dim pname As String, variantName As String, timestamp As Date

        Dim request As New Request(awinSettings.databaseName)
        Dim requestTrash As New Request(awinSettings.databaseName & "Trash")

        Dim p As Integer, v As Integer, t As Integer


        With TreeViewProjekte
            anzahlProjekte = .Nodes.Count

            For p = 1 To anzahlProjekte

                projektNode = .Nodes.Item(p - 1)
                pname = projektNode.Text

                If projektNode.Checked Then
                    ' Aktion auf allen Varianten und Timestamps 
                    ' Schleife über alle Varianten: 
                    ' lösche in Datenbank pname#vname
                    anzahlVarianten = projektNode.Nodes.Count

                    For v = 1 To anzahlVarianten

                        variantNode = projektNode.Nodes.Item(v - 1)
                        variantName = getVariantNameOf(variantNode.Text)
                        Call deleteCompleteProjectVariant(pname, variantName)

                    Next


                Else

                    anzahlVarianten = projektNode.Nodes.Count
                    For v = 1 To anzahlVarianten
                        variantNode = projektNode.Nodes.Item(v - 1)
                        variantName = getVariantNameOf(variantNode.Text)


                        If variantNode.Checked Then
                            ' Aktion auf allen Timestamps
                            ' lösche in Datenbank das Objekt mit DB-Namen pname#vname

                            Call deleteCompleteProjectVariant(pname, variantName)

                        Else

                            anzahlTimeStamps = variantNode.Nodes.Count
                            Dim first As Boolean = True

                            For t = 1 To anzahlTimeStamps
                                timeStampNode = variantNode.Nodes.Item(t - 1)

                                If timeStampNode.Checked Then
                                    ' Aktion auf diesem timestamp

                                    timestamp = CType(timeStampNode.Text, Date)
                                    Call deleteProjectVariantTimeStamp(pname, variantName, timestamp, first)


                                End If
                            Next
                        End If

                    Next
                End If

            Next


        End With


    End Sub


End Class