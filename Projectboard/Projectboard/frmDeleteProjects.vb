Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports System.Windows.Forms



Public Class frmDeleteProjects

    Public projekteInDB As New SortedList(Of String, clsProjekt)
    Public projektHistorien As New clsProjektDBInfos
    Public ProjZuLöschen As clsProjekt

    Private Sub DeleteButton_Click(sender As Object, e As EventArgs) Handles DeleteButton.Click

        'Call MsgBox("DeleteButton_Click")
    End Sub
    Private Sub frmDeleteProjects_FormClosed(sender As Object, e As EventArgs) Handles Me.FormClosed

        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left
        projektHistorien.clear()

    End Sub



    Private Sub frmDeleteProjects_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim treeview1 As New TreeView
        Dim node As TreeNode
        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(20)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = StartofCalendar
        Dim pname As String = ""
        Dim variantName As String = ""
     

        Dim deletedProj As Integer = 0
        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim anzElements As Integer
        Dim hproj As clsProjekt
        Dim schluessel As String = ""

        Dim request As New Request(awinSettings.databaseName)
        Dim requestTrash As New Request(awinSettings.databaseName & "Trash")

        projektHistorien.clear()

        ' Alle Projekte aus DB
        ' projekteInDB = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then    ' es sind Projekte selektiert

            If awinSelection.Count > 0 Then
                'selektierte Projekte ins Formular eintragen
                anzElements = awinSelection.Count
                Dim i As Integer
                For i = 1 To anzElements

                    singleShp = awinSelection.Item(i)
                    hproj = ShowProjekte.getProject(singleShp.Name)


                    schluessel = hproj.name & "#" & hproj.variantName

                    'If request.pingMongoDb() Then
                    '    ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                    '    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                    '                                                       storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                    'Else
                    '    Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                    '    projekthistorie.clear()
                    'End If

                    'If projekthistorie.Count > 0 Then
                    '    ' Aufbau der Listen 
                    '    projektHistorien.Add(projekthistorie)
                    'End If

                    projektHistorien.Add(schluessel, Date.MinValue) 'Platzhalter für die Projekthistorie

                    With TreeViewProjekte

                        .CheckBoxes = True

                        node = .Nodes.Add(hproj.name)
                        node.Nodes.Add(CType(Date.MinValue, String))    'Platzhalter für die Projekthistorie

                        'For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                        '    node.Nodes.Add(CType(kvp1.Key, String))
                        'Next kvp1

                        'If node.IsSelected Then
                        '    node.Expand()
                        'End If
                    End With
                Next i

            End If



        Else
            ' angezeigte Projekte ins Formular eintragen

            If ShowProjekte.Count > 0 Then

                With TreeViewProjekte

                    .CheckBoxes = True

                    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                        node = .Nodes.Add(kvp.Value.name)
                        hproj = kvp.Value
                        schluessel = hproj.name & "#" & hproj.variantName

                        projektHistorien.Add(schluessel, Date.MinValue) 'Platzhalter für die Projekthistorie in der Liste

                        node.Nodes.Add(CType(Date.MinValue, String))    'Platzhalter für die Projekthistorie im Formular


                        'If request.pingMongoDb() Then
                        '    ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        '    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                        '                                                       storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        'Else
                        '    Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                        '    projekthistorie.clear()
                        'End If

                        'If projekthistorie.Count > 0 Then
                        '    ' Aufbau der Listen 
                        '    projektHistorien.Add(projekthistorie)
                        'End If

                        'For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                        '    node.Nodes.Add(CType(kvp1.Key, String))
                        'Next kvp1

                        'If node.IsSelected Then
                        '    node.Expand()
                        'End If

                    Next kvp

                End With


            Else

                ' geladene Projekte ins Formular eintragen

                If AlleProjekte.Count > 0 Then

                    With TreeViewProjekte

                        .CheckBoxes = True

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte
                            node = .Nodes.Add(kvp.Value.name)
                            hproj = kvp.Value
                            schluessel = hproj.name & "#" & hproj.variantName

                            projektHistorien.Add(schluessel, Date.MinValue) 'Platzhalter für die Projekthistorie
                            node.Nodes.Add(CType(Date.MinValue, String))    'Platzhalter für die Projekthistorie im Formular

                            'If request.pingMongoDb() Then
                            '    ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                            '    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                            '                                                       storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                            'Else
                            '    Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                            '    projekthistorie.clear()
                            'End If

                            'If projekthistorie.Count > 0 Then
                            '    ' Aufbau der Listen 
                            '    projektHistorien.Add(projekthistorie)
                            'End If

                            'For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                            '    node.Nodes.Add(CType(kvp1.Key, String))
                            'Next kvp1

                            'If node.IsSelected Then
                            '    node.Expand()
                            'End If

                        Next kvp
                    End With



                Else
                    ' Alle Projekte aus DB

                    projekteInDB = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)

                    If projekteInDB.Count > 1 Then

                        With TreeViewProjekte

                            .CheckBoxes = True

                            For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteInDB
                                node = .Nodes.Add(kvp.Value.name)
                                hproj = kvp.Value
                                schluessel = hproj.name & "#" & hproj.variantName

                                projektHistorien.Add(schluessel, Date.MinValue) 'Platzhalter für die Projekthistorie
                                node.Nodes.Add(CType(Date.MinValue, String))    'Platzhalter für die Projekthistorie im Formular

                                'If request.pingMongoDb() Then
                                '    ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                                '    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                                '                                                       storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                                'Else
                                '    Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                                '    projekthistorie.clear()
                                'End If

                                'If projekthistorie.Count > 0 Then
                                '    ' Aufbau der Listen 
                                '    projektHistorien.Add(projekthistorie)
                                'End If

                                'For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                                '    node.Nodes.Add(CType(kvp1.Key, String))
                                'Next kvp1

                                'If node.IsSelected Then
                                '    node.Expand()
                                'End If


                            Next kvp
                        End With
                    Else
                        Call MsgBox(" keine Projekte in der Datenbank")
                    End If

                End If 'AlleProjekte

            End If 'showProjekte

        End If 'selektierte Projekte

    End Sub

    Private Sub TreeViewProjekte_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles TreeViewProjekte.AfterExpand

    End Sub
    Private Sub TreeViewProjekte_AfterSelect(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeViewProjekte.AfterSelect
      
    End Sub
    Private Sub TreeViewProjekte_AfterCheck(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeViewProjekte.AfterCheck
        Dim node As TreeNode
        Dim schluessel As String = ""
        Dim selCollection As SortedList(Of Date, String)
        Dim timeStamp As Date

        node = e.Node

        schluessel = CType(node.Text, String) & "#"

        If Not IsNothing(node.Parent) Then
            schluessel = CType(node.Parent.Text, String) & "#"
            Try
                selCollection = projektHistorien.getTimeStamps(schluessel)

                ' Löschen aus der projektHistorien-Liste
                ' projektHistorien.Remove(schluessel, CType(node.Text, Date))
                If node.Checked = True Then
                    ' Aufbau der Liste selectedToDelete
                    selectedToDelete.Add(schluessel, selCollection.ElementAt(node.Index).Key)
                Else
                    selectedToDelete.Remove(schluessel, selCollection.ElementAt(node.Index).Key)
                End If
            Catch ex As Exception

            End Try


        Else
            schluessel = CType(node.Text, String) & "#"
            Try


                If node.Checked = True Then

                    If node.IsExpanded = False Then
                        node.Expand()
                    End If

                    selCollection = projektHistorien.getTimeStamps(schluessel)

                    Dim i As Integer
                    For i = 1 To selCollection.Count
                        timeStamp = selCollection.ElementAt(i - 1).Key
                        selectedToDelete.Add(schluessel, timeStamp)

                        'Alle Unterknoten werden zum Löschen gecheckt
                        e.Node.Nodes(i - 1).Checked = True

                    Next i
                Else

                    selCollection = projektHistorien.getTimeStamps(schluessel)

                    Dim i As Integer
                    For i = 1 To selCollection.Count
                        timeStamp = selCollection.ElementAt(i - 1).Key
                        selectedToDelete.Remove(schluessel, timeStamp)

                        '' Check wird für alle Unterknoten entfernt
                        e.Node.Nodes(i - 1).Checked = False

                    Next i
                End If

            Catch ex As Exception

            End Try


        End If

    End Sub
    Private Sub TreeViewProjekte_AfterCollapse(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeViewProjekte.AfterCollapse
       
    End Sub
    Private Sub TreeViewProjekte_BeforeExpand(sender As Object, e As Windows.Forms.TreeViewCancelEventArgs) Handles TreeViewProjekte.BeforeExpand

        Dim request As New Request(awinSettings.databaseName)
        Dim node As New TreeNode
        Dim projName As String
        Dim variantName As String = ""
        Dim hliste As SortedList(Of Date, String)


        node = e.Node
        projName = node.Text

        hliste = projektHistorien.getTimeStamps(projName & "#" & variantName)

        If hliste.Count = 1 Then

            If hliste.ElementAt(0).Key = Date.MinValue Then

                If request.pingMongoDb() And request.projectNameAlreadyExists(projName, variantName) Then

                    Try
                        If Not projekthistorie Is Nothing Then
                            projekthistorie.clear()
                        End If

                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=projName, variantName:=variantName, _
                                                                         storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
                        'projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=projName, variantName:="", _
                        '                                                    storedEarliest:=StartofCalendar, storedLatest:=Date.Now)

                    Catch ex As Exception
                        projekthistorie = Nothing
                    End Try

                    If projekthistorie.Count > 0 Then

                        projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                        node.Nodes.Clear()  ' Löschen von Platzhalter

                        ' Aufbau der Listen 
                        projektHistorien.Add(projekthistorie)


                        ' Eintragen der zum Projekt gehörenden TimeStamps
                        For Each kvp1 As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
                            node.Nodes.Add(CType(kvp1.Value.timeStamp, String))
                        Next kvp1

                        If node.IsSelected Then
                            node.Expand()
                        End If
                    Else

                        If projekthistorie.Count = 0 Then
                            ' keine ProjektHistorie vorhanden
                            projektHistorien.Remove(projName & "#" & variantName, Date.MinValue) 'Platzhalter wieder entfernen
                            node.Nodes.Clear()  ' Löschen von Platzhalter
                        End If
                    End If

                Else
                    Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                End If

            End If
            ' es ist nichts zu machen, da die Historie zu diesem Projekt schon aus DB gelesen
        End If

    End Sub

    Private Sub AbbrechenButton_Click(sender As Object, e As EventArgs) Handles AbbrechenButton.Click
        'Call MsgBox("AbbrechenButton")
    End Sub


End Class