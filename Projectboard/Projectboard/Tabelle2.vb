
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel


Public Class Tabelle2

    Private columnStartData As Integer = 8
    Private columnEndData As Integer = 30
    Private columnRC As Integer = 5
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private columnName As Integer = 2


    Private Sub Tabelle2_ActivateEvent(Optional ByVal rcName As String = Nothing) Handles Me.ActivateEvent


        Application.DisplayFormulaBar = False

        Dim filterRange As Excel.Range
        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)


        ' jetzt den Schutz aufheben , falls einer definiert ist 
        If meWS.ProtectContents Then
            meWS.Unprotect(Password:="x")
        End If

        Try
            ' die Anzahl maximaler Zeilen bestimmen 
            With visboZustaende
                .meMaxZeile = CType(appInstance.ActiveSheet, Excel.Worksheet).UsedRange.Rows.Count
                .meColRC = CType(appInstance.ActiveSheet.Range("RoleCost"), Excel.Range).Column
                .meColSD = CType(appInstance.ActiveSheet.Range("StartData"), Excel.Range).Column
                .meColED = CType(appInstance.ActiveSheet.Range("EndData"), Excel.Range).Column
                .meColpName = 2
                columnRC = .meColRC
                columnStartData = .meColSD
                columnEndData = .meColED
            End With

        Catch ex As Exception
            Call MsgBox("Fehler in Laden des Sheets ...")
        End Try

        ' jetzt den AutoFilter setzen 
        Try
            ' der wird jetzt erst am Ende  gemacht 
            '' einen Select machen ...
            ''Try
            ''    'CType(CType(meWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
            ''    CType(CType(meWS, Excel.Worksheet).Cells(2, columnRC), Excel.Range).Select()
            ''Catch ex As Exception

            ''End Try


            With meWS
                filterRange = CType(.Range(.Cells(1, 1), .Cells(1, 6)), Excel.Range)
            End With

            ' jetzt die Autofilter aktivieren ... 
            If Not CType(meWS, Excel.Worksheet).AutoFilterMode = True Then
                'CType(meWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                CType(meWS, Excel.Worksheet).Rows(1).AutoFilter()

                ' jetzt überprüfen, ob nur eine bestimmte Rolle/Kostenart angezeigt, d.h gefiltert werden soll  
                If Not IsNothing(rcName) Then
                    CType(CType(meWS, Excel.Worksheet).Rows(1), Excel.Range).AutoFilter(Field:=visboZustaende.meColRC, Criteria1:=rcName)
                End If

            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Filtersetzen und Speichern" & vbLf & ex.Message)
        End Try

        Try
            If awinSettings.meEnableSorting Then

                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    ' braucht man nicht mehr - ist schon gemacht 
                    '.Unprotect("x")
                    .EnableSelection = XlEnableSelection.xlNoRestrictions
                End With
            Else
                With meWS
                    .Protect(Password:="x", UserInterfaceOnly:=True,
                             AllowFormattingCells:=True,
                             AllowFormattingColumns:=True,
                             AllowInsertingColumns:=False,
                             AllowInsertingRows:=True,
                             AllowDeletingColumns:=False,
                             AllowDeletingRows:=True,
                             AllowSorting:=True,
                             AllowFiltering:=True)
                    .EnableSelection = XlEnableSelection.xlUnlockedCells
                    .EnableAutoFilter = True
                End With
            End If


        Catch ex As Exception

        End Try

        ' jetzt soll geprüft werden, ob es sich um einen vglweise kleinen Bildschirm handelt - dann sollen 
        ' bestimmte Spaltengrößen verkleinert werden oder aber auch ausgeblendet werden .. oder Schriftgrößen verkleinert werden  

        ' das wird ja jetzt in der Defition der Windows gemacht ...
        'Try
        '    With Application.ActiveWindow
        '        .SplitColumn = columnRC + 2
        '        .SplitRow = 1
        '        .DisplayWorkbookTabs = False
        '        .GridlineColor = RGB(220, 220, 220)
        '        .FreezePanes = True
        '        '.DisplayHeadings = True
        '        .DisplayHeadings = False
        '    End With

        'Catch ex As Exception
        '    Call MsgBox("Fehler bei Activate Sheet Massen-Edit" & vbLf & ex.Message)
        'End Try

        ' tk 3.4.18 das wird jetzt im writeOnlineMassEditRC gemacht 
        'With meWS
        '    CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
        'End With

        If Not IsNothing(appInstance.ActiveCell) Then
            visboZustaende.oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If


        ' einen Select machen - nachdem Event Behandlung wieder true ist, dann werden project und lastprojectDB gesetzt ...
        Try
            'CType(CType(meWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
            ' jetzt auf die erste selektierbare Zeile gehen ... 
            Dim cz As Integer = 2
            Dim eof As Boolean = (cz > visboZustaende.meMaxZeile)

            Dim bedingung As Boolean = CBool(CType(meWS.Cells(cz, columnRC), Excel.Range).Locked = True) And Not eof

            Do While bedingung
                cz = cz + 1
                eof = (cz > visboZustaende.meMaxZeile)
                bedingung = CBool(CType(meWS.Cells(cz, columnRC), Excel.Range).Locked = True) And Not eof
            Loop

            If Not eof Then
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnRC), Excel.Range).Select()

                Dim pName As String = ""

                With visboZustaende

                    pName = CStr(CType(appInstance.ActiveSheet.Cells(cz, visboZustaende.meColpName), Excel.Range).Value)
                    If ShowProjekte.contains(pName) Then
                        .lastProject = ShowProjekte.getProject(pName)
                        .lastProjectDB = dbCacheProjekte.getProject(calcProjektKey(pName, .lastProject.variantName))
                    End If

                End With
            Else
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnRC), Excel.Range).Locked = False
            End If

            CType(CType(meWS, Excel.Worksheet).Cells(cz, columnRC), Excel.Range).Select()

        Catch ex As Exception

        End Try

        Application.EnableEvents = formerEE
        If Application.ScreenUpdating = False Then
            Application.ScreenUpdating = True
        End If



    End Sub

    Private Sub Tabelle2_BeforeDoubleClick(Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Me.BeforeDoubleClick

        ' ''Dim former_EE As Boolean = appInstance.EnableEvents

        ' ''appInstance.EnableEvents = True
        ' ''Dim currentCell As Excel.Range = Target

        ' ''Try
        ' ''    Dim frmMERoleCost As New frmMEhryRoleCost
        ' ''    Dim auslastungChanged As Boolean = False
        ' ''    Dim summenChanged As Boolean = False
        ' ''    ' muss extra überwacht werden, um das ProjectInfo1 Fenster auch immer zu aktualisieren
        ' ''    Dim kostenChanged As Boolean = False
        ' ''    Dim newStrValue As String = ""

        ' ''    Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
        ' ''    Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
        ' ''    Dim returnValue As DialogResult

        ' ''    If Target.Cells.Count = 1 Then

        ' ''        Dim zeile As Integer = Target.Row
        ' ''        Dim pName As String = CStr(meWS.Cells(zeile, visboZustaende.meColpName).value)
        ' ''        Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
        ' ''        Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
        ' ''        Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)
        ' ''        Dim phaseNameID As String = calcHryElemKey(phaseName, False)

        ' ''        Dim hproj As clsProjekt = Nothing
        ' ''        If Not IsNothing(pName) And pName <> "" Then
        ' ''            hproj = ShowProjekte.getProject(pName)
        ' ''        End If

        ' ''        Dim hPhase As clsPhase = hproj.getPhaseByID(phaseNameID)

        ' ''        Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
        ' ''        If Not IsNothing(curComment) Then
        ' ''            phaseNameID = curComment.Text
        ' ''        End If



        ' ''        If Target.Column = columnRC Then
        ' ''            ' es handelt sich um eine Rollen- oder Kosten-Änderung ...
        ' ''            ' Jetzt muss ein Formular mit den Rollen und Kosten im TreeView angezeigt werden
        ' ''            frmMERoleCost.pName = pName
        ' ''            frmMERoleCost.vName = vName
        ' ''            frmMERoleCost.phaseName = phaseName
        ' ''            frmMERoleCost.rcName = rcName
        ' ''            frmMERoleCost.phaseNameID = phaseNameID
        ' ''            frmMERoleCost.hproj = hproj

        ' ''            returnValue = frmMERoleCost.ShowDialog()

        ' ''            If returnValue = DialogResult.OK Then
        ' ''                ' eintragen der selektierten Rollen

        ' ''                If frmMERoleCost.ergItems.Count = 1 Then
        ' ''                    Dim hRCname As String = CStr(frmMERoleCost.ergItems.Item(1))

        ' ''                    ' jetzt den Schutz aufheben , falls einer definiert ist 
        ' ''                    If meWS.ProtectContents Then
        ' ''                        meWS.Unprotect(Password:="x")
        ' ''                    End If

        ' ''                    If rcName <> hRCname Then
        ' ''                        ' ausgewählte Rolle eintragn
        ' ''                        'CType(meWS.Cells(zeile, columnRC), Excel.Range).NumberFormat = Format("@")
        ' ''                        CType(meWS.Cells(zeile, columnRC), Excel.Range).Value = hRCname
        ' ''                        ' summe = 0 eintragen => es wird diese Rolle/Kosten in hproj eingetragen über change-event

        ' ''                        'CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).NumberFormat = Format("######0.0  ")
        ' ''                        CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = 0.0

        ' ''                        ' wenn es sich um eine Kostenart handelt, so wird ein Kommentar eingetragen
        ' ''                        If CostDefinitions.containsName(hRCname) Then

        ' ''                            CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).AddComment()
        ' ''                            With CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Comment
        ' ''                                .Visible = False
        ' ''                                If awinSettings.englishLanguage Then
        ' ''                                    .Text("Value in thousand €")
        ' ''                                Else
        ' ''                                    .Text(Text:="Angabe in T€")
        ' ''                                End If
        ' ''                                .Shape.ScaleHeight(0.6, Microsoft.Office.Core.MsoTriState.msoFalse)
        ' ''                            End With
        ' ''                        Else

        ' ''                            CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).ClearComments()
        ' ''                        End If

        ' ''                    End If
        ' ''                Else
        ' ''                    Dim i As Integer
        ' ''                    For i = 1 To frmMERoleCost.ergItems.Count

        ' ''                        If rcName = CStr(frmMERoleCost.ergItems(i)) Then
        ' ''                            ' aktuelle Rolle immer noch ausgewählt, muss aber nicht eingefügt werden, sondern nur alle anderen
        ' ''                        Else
        ' ''                            ' Zeile im MassenEdit-Tabelle einfügen und Namen einfügen
        ' ''                            Call massEditZeileEinfügen("")
        ' ''                            Dim hRCname As String = CStr(frmMERoleCost.ergItems.Item(i))

        ' ''                            If meWS.ProtectContents Then
        ' ''                                meWS.Unprotect(Password:="x")
        ' ''                            End If

        ' ''                            ' ausgewählte Rolle eintragn
        ' ''                            'CType(meWS.Cells(zeile, columnRC), Excel.Range).NumberFormat = Format("@")
        ' ''                            CType(meWS.Cells(zeile, columnRC), Excel.Range).Value = hRCname
        ' ''                            ' summe = 0 eintragen => es wird diese Rolle/Kosten in hproj eingetragen über change-event

        ' ''                            'CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).NumberFormat = Format("######0.0  ")
        ' ''                            CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = 0.0


        ' ''                            ' wenn es sich um eine Kostenart handelt, so wird ein Kommentar eingetragen
        ' ''                            If CostDefinitions.containsName(hRCname) Then
        ' ''                                ' jetzt den Schutz aufheben , falls einer definiert ist 

        ' ''                                CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).AddComment()
        ' ''                                With CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Comment
        ' ''                                    .Visible = False
        ' ''                                    If awinSettings.englishLanguage Then
        ' ''                                        .Text("Value in thousand €")
        ' ''                                    Else
        ' ''                                        .Text(Text:="Angabe in T€")
        ' ''                                    End If
        ' ''                                    .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
        ' ''                                End With
        ' ''                            Else

        ' ''                                '' ''CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Comment.Delete()
        ' ''                                CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).ClearComments()
        ' ''                            End If
        ' ''                        End If

        ' ''                    Next

        ' ''                End If
        ' ''                ' Blattschutz wieder setzen wie zuvor
        ' ''                With meWS
        ' ''                    .Protect(Password:="x", UserInterfaceOnly:=True, _
        ' ''                             AllowFormattingCells:=True, _
        ' ''                             AllowInsertingColumns:=False,
        ' ''                             AllowInsertingRows:=True, _
        ' ''                             AllowDeletingColumns:=False, _
        ' ''                             AllowDeletingRows:=True, _
        ' ''                             AllowSorting:=True, _
        ' ''                             AllowFiltering:=True)
        ' ''                    .EnableSelection = XlEnableSelection.xlUnlockedCells
        ' ''                    .EnableAutoFilter = True
        ' ''                End With
        ' ''            End If

        ' ''        End If

        ' ''    Else
        ' ''        Call MsgBox("bitte nur eine Zelle selektieren ...")
        ' ''        Target.Cells(1, 1).value = visboZustaende.oldValue
        ' ''    End If


        ' ''Catch ex As Exception

        ' ''    Call MsgBox("Fehler bei Massen-Edit, Ändern : " & vbLf & ex.Message)

        ' ''    ' Blattschutz wieder setzen wie zuvor
        ' ''    With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
        ' ''        .Protect(Password:="x", UserInterfaceOnly:=True, _
        ' ''                 AllowFormattingCells:=True, _
        ' ''                 AllowInsertingColumns:=False,
        ' ''                 AllowInsertingRows:=True, _
        ' ''                 AllowDeletingColumns:=False, _
        ' ''                 AllowDeletingRows:=True, _
        ' ''                 AllowSorting:=True, _
        ' ''                 AllowFiltering:=True)
        ' ''        .EnableSelection = XlEnableSelection.xlUnlockedCells
        ' ''        .EnableAutoFilter = True
        ' ''    End With

        ' ''End Try

        ' ''appInstance.EnableEvents = former_EE

    End Sub

    Private Sub Tabelle2_BeforeRightClick(Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Me.BeforeRightClick


        Dim former_EE As Boolean = appInstance.EnableEvents

        appInstance.EnableEvents = True
        Dim currentCell As Excel.Range = Target

        ' die Rechtsklick-Behandlung soll auf alle Fälle abgeschaltet werden 
        Cancel = True

        ' prüfen, ob sich das die selektierte Zelle in der Role-/Cost Spalte befindet 
        If Target.Column = columnRC Then

            Try
                Dim frmMERoleCost As New frmMEhryRoleCost
                Dim auslastungChanged As Boolean = False
                Dim summenChanged As Boolean = False
                ' muss extra überwacht werden, um das ProjectInfo1 Fenster auch immer zu aktualisieren
                Dim kostenChanged As Boolean = False
                Dim newStrValue As String = ""

                Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                Dim returnValue As DialogResult

                If Target.Cells.Count = 1 Then

                    Dim zeile As Integer = Target.Row
                    Dim pName As String = CStr(meWS.Cells(zeile, visboZustaende.meColpName).value)
                    Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
                    Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
                    Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)
                    Dim phaseNameID As String = calcHryElemKey(phaseName, False)

                    Dim hproj As clsProjekt = Nothing
                    If Not IsNothing(pName) And pName <> "" Then
                        hproj = ShowProjekte.getProject(pName)
                    End If

                    Dim hPhase As clsPhase = hproj.getPhaseByID(phaseNameID)

                    Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
                    If Not IsNothing(curComment) Then
                        phaseNameID = curComment.Text
                    End If



                    If Target.Column = columnRC Then
                        ' es handelt sich um eine Rollen- oder Kosten-Änderung ...
                        ' Jetzt muss ein Formular mit den Rollen und Kosten im TreeView angezeigt werden
                        frmMERoleCost.pName = pName
                        frmMERoleCost.vName = vName
                        frmMERoleCost.phaseName = phaseName
                        frmMERoleCost.rcName = rcName
                        frmMERoleCost.phaseNameID = phaseNameID
                        frmMERoleCost.hproj = hproj

                        returnValue = frmMERoleCost.ShowDialog()

                        If returnValue = DialogResult.OK Then
                            ' eintragen der selektierten Rollen

                            If frmMERoleCost.ergItems.Count = 1 Then
                                Dim hRCname As String = CStr(frmMERoleCost.ergItems.Item(1))

                                ' jetzt den Schutz aufheben , falls einer definiert ist 
                                If meWS.ProtectContents Then
                                    meWS.Unprotect(Password:="x")
                                End If
                                Dim rng As Excel.Range = CType(meWS.Cells(zeile, columnRC + 1), Excel.Range)
                                rng.ClearComments()


                                If rcName <> hRCname Then
                                    ' ausgewählte Rolle eintragn
                                    'CType(meWS.Cells(zeile, columnRC), Excel.Range).NumberFormat = Format("@")
                                    CType(meWS.Cells(zeile, columnRC), Excel.Range).Value = hRCname
                                    ' summe = 0 eintragen => es wird diese Rolle/Kosten in hproj eingetragen über change-event

                                    'CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).NumberFormat = Format("######0.0  ")
                                    If Not IsNumeric(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value) Then
                                        If CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = "" Then
                                            CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = 0.0
                                        End If
                                    End If

                                    ' wenn es sich um eine Kostenart handelt, so wird ein Kommentar eingetragen
                                    If CostDefinitions.containsName(hRCname) Then

                                        CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).AddComment()
                                        With CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Comment
                                            .Visible = False
                                            If awinSettings.englishLanguage Then
                                                .Text("Value in thousand €")
                                            Else
                                                .Text(Text:="Angabe in T€")
                                            End If
                                            .Shape.ScaleHeight(0.6, Microsoft.Office.Core.MsoTriState.msoFalse)
                                        End With
                                    Else

                                        '' jetzt den Schutz aufheben , falls einer definiert ist 
                                        'If meWS.ProtectContents Then
                                        '    meWS.Unprotect(Password:="x")
                                        'End If
                                        'Dim rng As Excel.Range = CType(meWS.Cells(zeile, columnRC + 1), Excel.Range)
                                        'rng.ClearComments()

                                    End If

                                End If
                            Else
                                Dim i As Integer
                                For i = 1 To frmMERoleCost.ergItems.Count

                                    If rcName = CStr(frmMERoleCost.ergItems(i)) Then
                                        ' aktuelle Rolle immer noch ausgewählt, muss aber nicht eingefügt werden, sondern nur alle anderen
                                    Else
                                        ' Zeile im MassenEdit-Tabelle einfügen und Namen einfügen
                                        ' es soll nur dann eine Zeile eingefügt werden, wenn bereits etwas für Rolle/Kostenart eingetragen ist 
                                        If i > 1 Or rcName <> "" Then
                                            Call massEditZeileEinfügen("")
                                            ' da in massEdit jetzt in der Zeile danach eins eingefügt wird, muss hier die zeile um eins erhöht werden ...
                                            zeile = zeile + 1
                                        End If

                                        Dim hRCname As String = CStr(frmMERoleCost.ergItems.Item(i))

                                        If meWS.ProtectContents Then
                                            meWS.Unprotect(Password:="x")
                                        End If
                                        Dim rng As Excel.Range = CType(meWS.Cells(zeile, columnRC + 1), Excel.Range)
                                        rng.ClearComments()


                                        ' ausgewählte Rolle eintragn
                                        'CType(meWS.Cells(zeile, columnRC), Excel.Range).NumberFormat = Format("@")
                                        CType(meWS.Cells(zeile, columnRC), Excel.Range).Value = hRCname
                                        ' summe = 0 eintragen => es wird diese Rolle/Kosten in hproj eingetragen über change-event

                                        'CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).NumberFormat = Format("######0.0  ")
                                        CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = 0.0


                                        ' wenn es sich um eine Kostenart handelt, so wird ein Kommentar eingetragen
                                        If CostDefinitions.containsName(hRCname) Then
                                            ' jetzt den Schutz aufheben , falls einer definiert ist 

                                            CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).AddComment()
                                            With CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Comment
                                                .Visible = False
                                                If awinSettings.englishLanguage Then
                                                    .Text("Value in thousand €")
                                                Else
                                                    .Text(Text:="Angabe in T€")
                                                End If
                                                .Shape.ScaleHeight(0.45, Microsoft.Office.Core.MsoTriState.msoFalse)
                                            End With
                                        Else

                                            ' '' ''CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Comment.Delete()
                                            ''CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).ClearComments()
                                            ' jetzt den Schutz aufheben , falls einer definiert ist 
                                            'If meWS.ProtectContents Then
                                            '    meWS.Unprotect(Password:="x")
                                            'End If
                                            'rng = CType(meWS.Cells(zeile, columnRC + 1), Excel.Range)
                                            'rng.ClearComments()

                                        End If
                                    End If

                                Next

                            End If
                            ' Blattschutz wieder setzen wie zuvor
                            'With meWS
                            '    .Protect(Password:="x", UserInterfaceOnly:=True, _
                            '             AllowFormattingCells:=True, _
                            '             AllowInsertingColumns:=False,
                            '             AllowInsertingRows:=True, _
                            '             AllowDeletingColumns:=False, _
                            '             AllowDeletingRows:=True, _
                            '             AllowSorting:=True, _
                            '             AllowFiltering:=True)
                            '    .EnableSelection = XlEnableSelection.xlUnlockedCells
                            '    .EnableAutoFilter = True
                            'End With

                            With meWS
                                .Protect(Password:="x", UserInterfaceOnly:=True,
                                    AllowFormattingCells:=True,
                                    AllowFormattingColumns:=True,
                                    AllowInsertingColumns:=False,
                                    AllowInsertingRows:=True,
                                    AllowDeletingColumns:=False,
                                    AllowDeletingRows:=True,
                                    AllowSorting:=True,
                                    AllowFiltering:=True)
                                .EnableSelection = XlEnableSelection.xlUnlockedCells
                                .EnableAutoFilter = True
                            End With
                            Cancel = True
                        End If

                    End If

                Else
                    Call MsgBox("bitte nur eine Zelle selektieren ...")
                    Target.Cells(1, 1).value = visboZustaende.oldValue
                End If


            Catch ex As Exception

                Call MsgBox("Fehler bei Massen-Edit, rightClick : " & vbLf & ex.Message)

                ' Blattschutz wieder setzen wie zuvor
                With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                    .Protect(Password:="x", UserInterfaceOnly:=True,
                             AllowFormattingCells:=True,
                             AllowFormattingColumns:=True,
                             AllowInsertingColumns:=False,
                             AllowInsertingRows:=True,
                             AllowDeletingColumns:=False,
                             AllowDeletingRows:=True,
                             AllowSorting:=True,
                             AllowFiltering:=True)
                    .EnableSelection = XlEnableSelection.xlUnlockedCells
                    .EnableAutoFilter = True
                End With

            End Try

        Else
            ' nichts weiter zu tun
        End If

        appInstance.EnableEvents = former_EE

    End Sub

    ''' <summary>
    ''' wird aufgerufen, sobald sich der Wert in einer Zelle verändert hat ...
    ''' entweder nachdem eine Dropbox Selection getroffen wurde oder eine Eingabe duch Pfeiltaste / Eingabe beendet wurde
    ''' 
    ''' </summary>
    ''' <param name="Target"></param>
    ''' <remarks></remarks>
    Private Sub Tabelle2_Change(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.Change

        ' damit nicht eine immerwährende Event Orgie durch Änderung in den Zellen abgeht ...
        appInstance.EnableEvents = False
        Dim currentCell As Excel.Range = Target

        Try
            Dim auslastungChanged As Boolean = False
            Dim summenChanged As Boolean = False
            ' muss extra überwacht werden, um das ProjectInfo1 Fenster auch immer zu aktualisieren
            Dim kostenChanged As Boolean = False
            Dim newStrValue As String = ""

            Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
            Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

            If Target.Cells.Count = 1 Then

                Dim roleCostNames As New Collection

                Dim zeile As Integer = Target.Row
                Dim pName As String = CStr(meWS.Cells(zeile, visboZustaende.meColpName).value)
                Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
                Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
                Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)
                Dim phaseNameID As String = calcHryElemKey(phaseName, False)
                Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
                If Not IsNothing(curComment) Then
                    phaseNameID = curComment.Text
                End If

                If Target.Column = columnRC Then
                    ' es handelt sich um eine Rollen- oder Kosten-Änderung ...


                    newStrValue = CStr(Target.Cells(1, 1).value)
                    If isValidRCChange(newStrValue, visboZustaende.oldValue) Then
                        ' es ist eine gültige Änderung, das heisst es wurde eine Rolle in eine andere gewechselt , oder 
                        ' eine Kostenart in eine andere; Kategorie-übergreifende Wechsel sind nicht erlaubt 

                        ' jetzt muss noch geprüft werden, ob auch keine Duplikate vorkommen: zu einem Projekt dürfen z.Bsp keine 
                        ' 2 Zeilen existieren mit jeweils der gleichen Rolle oder Kostenart ...
                        If noDuplicatesInSheet(pName, phaseNameID, newStrValue, zeile) Then

                            Dim hproj As clsProjekt = ShowProjekte.getProject(pName)

                            ' jetzt werden die Validation-Strings für alles, alleRollen, alleKosten und die einzelnen SammelRollen aufgebaut 
                            Dim validationStrings As SortedList(Of String, String) = createMassEditRcValidations()
                            Dim anzahlRollen As Integer = RoleDefinitions.Count
                            Dim rcValidation() As String
                            ' in rcValidation(0) steht der Name "alleKosten" für den Validation-String für alle Kosten
                            ' in rcValidation(i) steht der Name des Validation-String für Rolle mit UID i 
                            ReDim rcValidation(anzahlRollen + 1)

                            rcValidation(0) = "alleKosten"
                            rcValidation(anzahlRollen + 1) = "alles"

                            For i As Integer = 1 To anzahlRollen
                                Dim tmprole As clsRollenDefinition = RoleDefinitions.getRoledef(i)
                                If tmprole.isCombinedRole Then
                                    rcValidation(i) = tmprole.name
                                Else
                                    Dim parentrole As clsRollenDefinition = RoleDefinitions.getParentRoleOf(tmprole.UID)
                                    If IsNothing(parentrole) Then
                                        rcValidation(i) = "alleRollen"
                                    Else
                                        rcValidation(i) = parentrole.name
                                    End If

                                End If
                            Next
                            ' Ende Preparation für Validierungs-Strings


                            If Not IsNothing(hproj) Then
                                Dim cPhase As clsPhase = hproj.getPhaseByID(phaseNameID)

                                If Not IsNothing(cPhase) Then
                                    If RoleDefinitions.containsName(newStrValue) Then
                                        ' es handelt sich um eine Rollen-Änderung
                                        Dim newRoleID As Integer = RoleDefinitions.getRoledef(newStrValue).UID
                                        If visboZustaende.oldValue.Length > 0 And visboZustaende.oldValue.Trim <> newStrValue.Trim Then
                                            ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                            Try
                                                auslastungChanged = True
                                                Dim cRole As clsRolle = cPhase.getRole(visboZustaende.oldValue)
                                                If IsNothing(cRole) Then
                                                Else
                                                    hproj.rcLists.removeRP(cRole.RollenTyp, cPhase.nameID)
                                                    cRole.RollenTyp = newRoleID
                                                    hproj.rcLists.addRP(newRoleID, cPhase.nameID)
                                                End If

                                            Catch ex As Exception
                                                visboZustaende.oldValue = ""
                                                ' in diesem Fall wurde nur von einer noch nicht belegten Rolle auf eine 
                                                ' andere nicht belegte gewechselt 
                                            End Try

                                        Else
                                            ' es kam eine neue Rolle hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
                                            ' muss an dieser Stelle nur die  gar nichts gemacht werden ..
                                            ' es sollen aber gleich die Auslastungs-Werte aktualisiert werden ...
                                            auslastungChanged = True
                                        End If

                                        ' jetzt für die Zelle die Validation neu bestimmen, dazu muss aber der Blattschutz aufgehoben sein ...  

                                        If Not awinSettings.meEnableSorting Then
                                            ' es muss der Blattschutz aufgehoben werden, nachher wieder mit diesen Einstellungen aktiviert werden ...
                                            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                                .Unprotect(Password:="x")
                                            End With
                                        End If

                                        With currentCell

                                            Try
                                                If Not IsNothing(.Validation) Then
                                                    .Validation.Delete()
                                                End If
                                                ' jetzt wird die ValidationList aufgebaut 
                                                Dim tmpVal As String = validationStrings.Item(rcValidation(newRoleID))

                                                '' ur: 28.09.2017
                                                ''.Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                ''                               Formula1:=tmpVal)

                                                ' wenn es sich um die Projekt-Phase handelt
                                                If phaseNameID = rootPhaseName Then
                                                    tmpVal = tmpVal & ";" &
                                                                validationStrings.Item(rcValidation(0))
                                                    Call updateEmptyRcCellValidations(pName, tmpVal)
                                                End If

                                            Catch ex As Exception

                                            End Try
                                        End With

                                        If Not awinSettings.meEnableSorting Then
                                            ' es muss der Blattschutz aufgehoben werden, nachher wieder mit diesen Einstellungen aktiviert werden ...
                                            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                                .Protect(Password:="x", UserInterfaceOnly:=True,
                                                         AllowFormattingCells:=True,
                                                         AllowFormattingColumns:=True,
                                                         AllowInsertingColumns:=False,
                                                         AllowInsertingRows:=True,
                                                         AllowDeletingColumns:=False,
                                                         AllowDeletingRows:=True,
                                                         AllowSorting:=True,
                                                         AllowFiltering:=True)
                                                .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
                                                .EnableAutoFilter = True
                                            End With
                                        End If

                                        ' jetzt die Rollen bestimmen, die neu berechnet werden müssen ... 
                                        roleCostNames = RoleDefinitions.getSummaryRoles(newStrValue)
                                        If Not roleCostNames.Contains(newStrValue) Then
                                            roleCostNames.Add(newStrValue, newStrValue)
                                        End If

                                        If visboZustaende.oldValue.Length > 0 Then
                                            If Not roleCostNames.Contains(visboZustaende.oldValue) Then
                                                roleCostNames.Add(visboZustaende.oldValue, visboZustaende.oldValue)
                                            End If
                                            Dim tmpSummaryNames As Collection = RoleDefinitions.getSummaryRoles(visboZustaende.oldValue)
                                            For sr As Integer = 1 To tmpSummaryNames.Count
                                                Dim srName As String = CStr(tmpSummaryNames.Item(sr))
                                                If Not roleCostNames.Contains(srName) Then
                                                    roleCostNames.Add(srName, srName)
                                                End If
                                            Next
                                        End If
                                    Else
                                        ' es handelt sich um eine Kostenart Änderung 
                                        If visboZustaende.oldValue.Length > 0 And visboZustaende.oldValue.Trim <> newStrValue.Trim Then
                                            ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                            Dim newCostID As Integer = CostDefinitions.getCostdef(newStrValue).UID
                                            Dim cCost As clsKostenart = cPhase.getCost(visboZustaende.oldValue)
                                            If IsNothing(cCost) Then
                                            Else
                                                hproj.rcLists.removeCP(cCost.KostenTyp, cPhase.nameID)
                                                cCost.KostenTyp = newCostID
                                                hproj.rcLists.addCP(newCostID, cPhase.nameID)
                                            End If
                                            kostenChanged = True
                                        Else
                                            ' es kam eine neue Kostenart hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
                                            ' muss an dieser Stelle noch gar nichts gemacht werden ..
                                        End If
                                    End If



                                Else
                                    Call MsgBox("Projekt-Phase kann nicht bestimmt werden: " & pName & ", " & phaseName)
                                End If
                            Else
                                Call MsgBox("Projekt kann nicht bestimmt werden: " & pName)
                            End If



                        Else
                            Call MsgBox("keine Doppelbelegung innerhalb einer Projektphase erlaubt ... ")
                            Target.Cells(1, 1).value = visboZustaende.oldValue

                            If visboZustaende.oldValue = "" Or IsNothing(visboZustaende.oldValue) Then
                                ' Zeile löschen mit Doppelbelegung
                                Call massEditZeileLoeschen("")

                            ElseIf RoleDefinitions.containsName(visboZustaende.oldValue) Then
                                Target.ClearComments()

                            End If


                        End If



                    Else
                        Call MsgBox("bitte nur innerhalb Rollen bzw. innerhalb Kostenarten wechseln !")
                        Target.Cells(1, 1).value = visboZustaende.oldValue
                    End If


                ElseIf Target.Column = columnRC + 1 Then
                    ' es handelt sich um eine Summenänderung
                    Dim newDblValue As Double
                    Dim difference As Double
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    Dim ok As Boolean = False
                    Dim isRole As Boolean
                    Dim uid As Integer

                    If RoleDefinitions.containsName(rcName) Then
                        isRole = True
                        uid = RoleDefinitions.getRoledef(rcName).UID
                        ok = True
                    ElseIf CostDefinitions.containsName(rcName) Then
                        isRole = False
                        uid = CostDefinitions.getCostdef(rcName).UID
                        ok = True
                    Else
                        Call MsgBox("bitte erst eine Rolle oder Kostenart auswählen !")
                        Target.Cells(1, 1).value = visboZustaende.oldValue
                    End If

                    If ok Then

                        If inputIsAcknowledged(Target, newDblValue, difference) Then

                            If Not IsNothing(hproj) Then
                                Dim cPhase As clsPhase = hproj.getPhaseByID(phaseNameID)

                                If Not IsNothing(cPhase) Then

                                    Dim phStart As Integer = hproj.Start + cPhase.relStart - 1
                                    Dim phEnde As Integer = hproj.Start + cPhase.relEnde - 1

                                    Dim ixZeitraum As Integer
                                    Dim ix As Integer
                                    Dim breite As Integer
                                    Call awinIntersectZeitraum(phStart, phEnde, ixZeitraum, ix, breite)

                                    Dim vSum As Double()
                                    ReDim vSum(0)
                                    vSum(0) = newDblValue
                                    Dim xStartDate As Date
                                    Dim xEndDate As Date

                                    If ix = 0 Then
                                        xStartDate = cPhase.getStartDate
                                    Else
                                        xStartDate = cPhase.getStartDate.AddDays(-1 * (cPhase.getStartDate.Day - 1)).AddMonths(ix)
                                    End If

                                    xEndDate = xStartDate.AddDays(-1 * (xStartDate.Day - 1)).AddMonths(breite).AddDays(-1)

                                    If DateDiff(DateInterval.Day, cPhase.getEndDate, xEndDate) > 0 Then
                                        xEndDate = cPhase.getEndDate
                                    End If

                                    Dim xValues() As Double = cPhase.berechneBedarfeNew(xStartDate,
                                                                                        xEndDate, vSum, 1)

                                    If isRole Then

                                        ' erstmal überprüfen, ob awinsettings.autoreduce = true 
                                        Dim parentRoleSum As Double = -1
                                        If awinSettings.meAutoReduce Then
                                            Call autoReduceRowOfParentRole(Target.Row, Target.Column, newDblValue, difference,
                                                                           hproj, cPhase, rcName)

                                            ' durch autoReduce kann der newDblValue verändert sein
                                            vSum(0) = newDblValue
                                            xValues = cPhase.berechneBedarfeNew(xStartDate,
                                                                                       xEndDate, vSum, 1)

                                        End If

                                        ' jetzt muss die Rolle aktualisiert werden ...
                                        Dim tmpRole As clsRolle = cPhase.getRole(rcName)

                                        If IsNothing(tmpRole) Then
                                            tmpRole = New clsRolle(phEnde - phStart)

                                            With tmpRole
                                                .RollenTyp = uid
                                            End With
                                            With cPhase
                                                .addRole(tmpRole)
                                            End With
                                        End If

                                        If tmpRole.Xwerte.Length <> xValues.Length Then
                                            For lx As Integer = 0 To breite - 1
                                                tmpRole.Xwerte(lx + ix) = xValues(lx)
                                            Next
                                        Else
                                            For i As Integer = 0 To tmpRole.Xwerte.Length - 1
                                                tmpRole.Xwerte(i) = xValues(i)
                                            Next
                                        End If

                                        auslastungChanged = True

                                        ' jetzt muss die Excel Zeile geschreiben werden - dort wird auch der auslastungs-Array aktualisiert 
                                        Call aktualisiereRoleCostInSheet(Target.Row, rcName, isRole,
                                                                     visboZustaende.meColSD, showRangeLeft, showRangeRight,
                                                                     phStart, phEnde, xValues)


                                    Else
                                        ' es handelt sich um eine Kostenart 
                                        Dim tmpCost As clsKostenart = cPhase.getCost(rcName)

                                        If IsNothing(tmpCost) Then
                                            tmpCost = New clsKostenart(phEnde - phStart)

                                            With tmpCost
                                                .KostenTyp = uid
                                            End With
                                            With cPhase
                                                .AddCost(tmpCost)
                                            End With
                                        End If

                                        If tmpCost.Xwerte.Length <> xValues.Length Then
                                            For lx As Integer = 0 To breite - 1
                                                tmpCost.Xwerte(lx + ix) = xValues(lx)
                                            Next
                                        Else
                                            For i As Integer = 0 To tmpCost.Xwerte.Length - 1
                                                tmpCost.Xwerte(i) = xValues(i)
                                            Next
                                        End If

                                        kostenChanged = True
                                        ' jetzt muss die Excel Zeile geschreiben werden 
                                        Call aktualisiereRoleCostInSheet(Target.Row, rcName, isRole,
                                                                     visboZustaende.meColSD, showRangeLeft, showRangeRight,
                                                                     phStart, phEnde, xValues)

                                    End If


                                End If
                            End If

                        Else
                            ' nichts tun 
                        End If

                    End If



                Else

                    ' es handelt sich um eine Datenänderung
                    Dim newDblValue As Double
                    Dim difference As Double

                    ' zu welcher / welchen Sammelrollen gehört die ausgewählte Rolle ? 
                    Dim sammelRollenName As String = ""
                    Dim zeileSammelRolle As Integer = 0
                    Dim isRole As Boolean

                    If RoleDefinitions.containsName(rcName) Then
                        isRole = True
                        ' hier muss jetzt bestimmt werden, wo die zugehörige Sammelrolle steht ... 
                    End If

                    If isRole Or CostDefinitions.containsName(rcName) Then
                        ' hier ist etwas gültiges vorhanden .. es kann also weitergemacht werden 

                        Try
                            If IsNothing(Target.Cells(1, 1).value) Then
                                newDblValue = 0.0
                            ElseIf IsNumeric(Target.Cells(1, 1).value) Then
                                newDblValue = CDbl(Target.Cells(1, 1).value)
                            Else
                                newDblValue = 0.0
                            End If
                        Catch ex As Exception
                            newDblValue = 0.0
                        End Try

                        Try
                            If IsNothing(visboZustaende.oldValue) Then
                                difference = newDblValue
                                visboZustaende.oldValue = "0"
                            ElseIf visboZustaende.oldValue = "" Then
                                difference = newDblValue
                                visboZustaende.oldValue = "0"
                            Else
                                difference = newDblValue - CDbl(visboZustaende.oldValue)
                            End If
                        Catch ex As Exception
                            difference = newDblValue
                            visboZustaende.oldValue = "0"
                        End Try

                        Dim monthCol As Integer = showRangeLeft + CInt(((Target.Column - columnStartData) / 2))

                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)

                        If Not IsNothing(hproj) Then
                            Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)

                            If Not IsNothing(cphase) Then

                                Dim xWerteIndex As Integer = monthCol - getColumnOfDate(cphase.getStartDate)
                                Dim xWerte() As Double

                                If isRole Then
                                    ' es handelt sich um eine gültige Rolle

                                    If awinSettings.meAutoReduce Then

                                        Call autoReduceCellOfParentRole(Target.Row, Target.Column, newDblValue,
                                                                  hproj, cphase, rcName, xWerteIndex, difference, summenChanged)

                                    End If

                                    ' es muss einfach die Rolle hinzugefügt bzw. die Werte abgeändert werden 
                                    Dim tmpRole As clsRolle = cphase.getRole(rcName)

                                    If IsNothing(tmpRole) Then
                                        ' die Rolle muss neu angelegt und der Phase hinzugefügt werden  

                                        tmpRole = New clsRolle(cphase.relEnde - cphase.relStart)
                                        tmpRole.RollenTyp = RoleDefinitions.getRoledef(rcName).UID

                                        Call cphase.addRole(tmpRole)

                                    End If

                                    ' der Monatswert muss geändert werden 
                                    xWerte = tmpRole.Xwerte
                                    If xWerteIndex >= 0 And xWerteIndex <= xWerte.Length - 1 Then
                                        If xWerte(xWerteIndex) <> newDblValue Then
                                            xWerte(xWerteIndex) = newDblValue
                                            summenChanged = True
                                        End If
                                    Else
                                        Call MsgBox("Fehler in Übernahme Daten-Wert ...")
                                    End If

                                    'tmpSum = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)
                                    'tmpSum = tmpSum + difference
                                    'CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = tmpSum

                                    ' bestimmt zu welchen Rollen die Auslastungs-Werte neu berechnet werden müssen ..
                                    roleCostNames = RoleDefinitions.getSummaryRoles(rcName)
                                    If Not roleCostNames.Contains(rcName) Then
                                        roleCostNames.Add(rcName, rcName)
                                    End If

                                    ' ur: 24.11.2017: Neuberechnung der Auslastung soll hier angestoßen werden, da Veränderung an Rolle in einem Monat mit entsprechenden Reduktion in Sammelrolle
                                    '
                                    'If difference <> 0 Then
                                    '    auslastungChanged = True
                                    'End If

                                    auslastungChanged = True


                                Else
                                    ' es handelt sich um eine gültige Kostenart - weiter oben wurde ja schon bestimmt, dass es entweder eine 
                                    ' gültige Rolle oder Kotenart ist 

                                    ' es muss einfach die Kostenart hinzugefügt bzw. die Werte abgeändert werden 
                                    Dim tmpCost As clsKostenart = cphase.getCost(rcName)

                                    If IsNothing(tmpCost) Then
                                        ' die Kostenart muss neu angelegt und der Phase hinzugefügt werden  

                                        tmpCost = New clsKostenart(cphase.relEnde - cphase.relStart)
                                        tmpCost.KostenTyp = CostDefinitions.getCostdef(rcName).UID

                                        Call cphase.AddCost(tmpCost)

                                        kostenChanged = True
                                    End If

                                    ' der Monatswert muss geändert werden 
                                    xWerte = tmpCost.Xwerte
                                    If xWerteIndex >= 0 And xWerteIndex <= xWerte.Length - 1 Then
                                        xWerte(xWerteIndex) = newDblValue
                                        summenChanged = True
                                    Else
                                        Call MsgBox("Fehler in Übernahme Daten-Wert ...")
                                    End If

                                    If Not roleCostNames.Contains(rcName) Then
                                        roleCostNames.Add(rcName, rcName)
                                    End If

                                End If
                            Else
                                Call MsgBox("Projekt-Phase existiert nicht: " & pName & ", " & phaseName)
                            End If
                        Else
                            Call MsgBox("Projekt existiert nicht: " & pName)
                        End If


                    Else
                        Call MsgBox("bitte erst eine Rolle oder Kostenart auswählen !")
                        Target.Cells(1, 1).value = visboZustaende.oldValue
                    End If



                End If


                If auslastungChanged And awinSettings.meExtendedColumnsView Then
                    Call updateMassEditAuslastungsValues(showRangeLeft, showRangeRight, roleCostNames)
                End If

                ' das Folgende ist eigentlich eine Test Routine , die normalerweise gar nicht nötig ist 
                ' aber für Testzwecke gut geeignet ist ...

                'Dim testValue1 As Double = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)
                If summenChanged Then
                    Call updateMassEditSummenValues(pName, phaseNameID, showRangeLeft, showRangeRight, roleCostNames)
                End If
                'Dim testValue2 As Double = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)

                'If testValue1 <> testValue2 Then
                '    Call MsgBox("Unterschiede: " & testValue1 & ", " & testValue2)
                'End If

                If Not IsNothing(Target.Cells(1, 1).value) Then
                    visboZustaende.oldValue = CStr(Target.Cells(1, 1).value)
                Else
                    visboZustaende.oldValue = ""
                End If

                ' aktualisieren der Charts 
                Try

                    If auslastungChanged Or summenChanged Or kostenChanged Then
                        If Not IsNothing(formProjectInfo1) Then
                            Call updateProjectInfo1(visboZustaende.lastProject, visboZustaende.lastProjectDB)
                        End If
                        Call aktualisiereCharts(visboZustaende.lastProject, True)
                        Call awinNeuZeichnenDiagramme(typus:=6, roleCost:=rcName)
                    End If

                Catch ex As Exception

                End Try

            Else
                Call MsgBox("bitte nur eine Zelle selektieren ...")
                Target.Cells(1, 1).value = visboZustaende.oldValue
            End If


        Catch ex As Exception
            Call MsgBox("Fehler bei Massen-Edit, Ändern : " & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True
    End Sub

    ''' <summary>
    ''' reduziert / erhöht den Sammelrollen Wert entsprechend der Änderung im Feld 
    ''' reduziert wird in der Phase als auch im Sheet 
    ''' ggf wird auch der newValue und difference neu bestimmt, deswegen Übergabe byref ...  
    ''' </summary>
    ''' <param name="targetRow"></param>
    ''' <param name="targetColumn"></param>
    ''' <param name="newValue"></param>
    ''' <param name="hproj"></param>
    ''' <param name="cPhase"></param>
    ''' <param name="roleName"></param>
    ''' <param name="xWerteIndex"></param>
    ''' <param name="difference"></param>
    ''' <param name="summenChanged"></param>
    ''' <remarks></remarks>
    Private Sub autoReduceCellOfParentRole(ByVal targetRow As Integer, ByVal targetColumn As Integer, ByRef newValue As Double,
                                         ByVal hproj As clsProjekt, ByVal cPhase As clsPhase, ByVal roleName As String,
                                         ByVal xWerteIndex As Integer, ByRef difference As Double,
                                         ByRef summenChanged As Boolean)

        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

        Dim pName As String = hproj.name
        Dim phaseNameID As String = cPhase.nameID

        Dim zeileOFSummaryRole As Integer = findeSammelRollenZeile(pName, phaseNameID, roleName)

        If zeileOFSummaryRole >= 2 And zeileOFSummaryRole <= visboZustaende.meMaxZeile Then

            Dim parentRoleName As String = CStr(meWS.Cells(zeileOFSummaryRole, columnRC).value)
            Dim parentPhaseName As String = CStr(meWS.Cells(zeileOFSummaryRole, 4).value)
            Dim parentPhaseNameID As String = calcHryElemKey(parentPhaseName, False)
            Dim parentComment As Excel.Comment = CType(meWS.Cells(zeileOFSummaryRole, 4), Excel.Range).Comment
            Dim xWerte() As Double

            If Not IsNothing(parentComment) Then
                phaseNameID = parentComment.Text
            End If

            Dim cParentPhase As clsPhase
            If parentPhaseNameID = phaseNameID Then
                cParentPhase = cPhase
            Else
                cParentPhase = hproj.getPhaseByID(parentPhaseNameID)
            End If

            ' das ist der Wert, um den der Index für die Parentphase korrigiert werden muss, da ja 
            ' die RootPhase wesentlich weiter links anfangen kann als die cphase
            ' es ist sicher gestellt, dass nur in zulässigen Wertebereichen aktualisiert wird 
            Dim offset As Integer = cPhase.relStart - cParentPhase.relStart

            ' jetzt muss in der Sammel-Rolle aktualisiert werden 
            Dim parentRole As clsRolle = Nothing
            Try
                parentRole = cParentPhase.getRole(parentRoleName)
            Catch ex As Exception

            End Try


            If IsNothing(parentRole) Then
                ' nichts tun 
            Else
                ' der Monatswert muss in der parentRole geändert werden 
                xWerte = parentRole.Xwerte
                If xWerteIndex + offset >= 0 And xWerteIndex + offset <= xWerte.Length - 1 Then
                    Dim alterWert As Double = xWerte(xWerteIndex + offset)
                    Dim savDifferenz As Double = difference
                    Dim sumRoleSum As Double = 0
                    Dim verteilungMöglich As Boolean = False
                    Dim msgResult As MsgBoxResult = MsgBoxResult.No

                    ' Test, ob es überhaupt möglich ist den eingegebenen Wert bei der Sammelrolle abzuziehen
                    ' ''For i As Integer = 0 To xWerte.Length - 1 - xWerteIndex - offset
                    ' ''    sumRoleSum = sumRoleSum + xWerte(xWerteIndex + offset + i)
                    ' ''Next

                    sumRoleSum = xWerte.Sum

                    If sumRoleSum >= difference Then
                        ' das darf aber nur gelöscht werden, wenn die Phase komplett im showrangeleft / showrangeright liegt 
                        If phaseWithinTimeFrame(hproj.Start, cPhase.relStart, cPhase.relEnde,
                                                 showRangeLeft, showRangeRight, True) Then

                            verteilungMöglich = True

                            ''Call MsgBox("die Phase wird nicht vollständig angezeigt - deshalb kann die Rolle " & rcName & vbLf & _
                            ''            " nicht gelöscht werden ...")
                            ''ok = False
                        Else
                            verteilungMöglich = False
                        End If

                    End If

                    If Not verteilungMöglich Or Not awinSettings.meDontAskWhenAutoReduce Then

                        xWerte(xWerteIndex + offset) = xWerte(xWerteIndex + offset) - difference


                        If xWerte(xWerteIndex + offset) < 0 Then

                            ' jetzt muss der newValue entsprechend geändert werden 
                            ' plus, weil xWerte(..) < 0 
                            newValue = newValue + xWerte(xWerteIndex + offset)

                            ' jetzt muss eine Meldung erfolgen ... 

                            Call MsgBox("AutoReduce kann die zugehörige Sammelrolle nicht auf negative Werte reduzieren" & vbLf &
                                        "oder die Phase wird nicht vollständig dargestellt" & vbLf &
                                        "Der Wert wird deshalb von " & CType(meWS.Cells(targetRow, targetColumn), Excel.Range).Value &
                                        " auf " & newValue & " korrigiert.")

                            difference = -xWerte(xWerteIndex + offset)


                            ' jetzt muss der newValue in das Feld geschrieben werden 
                            CType(meWS.Cells(targetRow, targetColumn), Excel.Range).Value = newValue

                            ' die Monatszahl und dann die Summe updaten ... 
                            xWerte(xWerteIndex + offset) = 0
                            CType(meWS.Cells(zeileOFSummaryRole, targetColumn), Excel.Range).Value = xWerte(xWerteIndex + offset)

                            CType(meWS.Cells(zeileOFSummaryRole, 6), Excel.Range).Value = xWerte.Sum

                        Else

                            difference = 0
                        End If


                    Else


                        For i As Integer = 0 To xWerte.Length - 1 - xWerteIndex - offset

                            xWerte(xWerteIndex + offset + i) = xWerte(xWerteIndex + offset + i) - difference


                            If xWerte(xWerteIndex + offset + i) < 0 Then

                                ' ''If i < 1 Then


                                ' ''    msgResult = MsgBox("AutoReduce kann die zugehörige Sammelrolle in diesem Monat nicht auf negative Werte reduzieren" & vbLf & _
                                ' ''                       "Soll der Wert deshalb von " & CType(meWS.Cells(targetRow, targetColumn), Excel.Range).Value & _
                                ' ''                       " auf " & newValue + xWerte(xWerteIndex + offset + i) & " korrigiert werden? (Ja)" & vbLf & _
                                ' ''                       "oder in den nächsten Monaten reduziert werden? (Nein)", MsgBoxStyle.YesNo)
                                ' ''    'msgResult = MsgBoxResult.Yes
                                ' ''    ''Else
                                ' ''    ''    ' es soll in den nächsten Monaten reduziert werden
                                ' ''    ''    msgResult = MsgBoxResult.No
                                ' ''End If



                                ' ''If msgResult = MsgBoxResult.Yes Then

                                ' ''    ' jetzt muss der newValue entsprechend geändert werden 
                                ' ''    ' plus, weil xWerte(..) < 0 
                                ' ''    newValue = newValue + xWerte(xWerteIndex + offset + i)

                                ' ''    ' jetzt muss der newDblValue in das Feld geschrieben werden 
                                ' ''    CType(meWS.Cells(targetRow, targetColumn + 2 * i), Excel.Range).Value = newValue
                                ' ''    ' die Monatszahl und dann die Summe updaten ... 
                                ' ''    CType(meWS.Cells(zeileOFSummaryRole, targetColumn + 2 * i), Excel.Range).Value = 0

                                ' ''    difference = -xWerte(xWerteIndex + offset + i)
                                ' ''    Exit For
                                ' ''Else

                                ' zu wenig abgezogen, wird in nächstem Monat abgezogen
                                Dim zuwenig As Double = -xWerte(xWerteIndex + offset + i)


                                ' bestimmen der neuen Differenz 
                                'ur:4.10..2017: difference = newValue - CDbl(visboZustaende.oldValue)
                                difference = zuwenig

                                ' ur: 4.10.2017: hier muss die Verteilung von  "difference"  stattfinden

                                xWerte(xWerteIndex + offset + i) = 0
                                ' ''End If

                            Else

                                difference = 0

                            End If


                            'die Monatszahl und dann die Summe updaten ... 
                            '' ''Dim testdbl As Double = xWerte(xWerteIndex + offset + i)
                            '' ''Call MsgBox(" testdbl = " & testdbl.ToString)
                            CType(meWS.Cells(zeileOFSummaryRole, targetColumn + 2 * i), Excel.Range).Value = xWerte(xWerteIndex + offset + i)

                        Next i

                        ' nun noch die Werte vor Beginn aktuellem Monat betrachten, sofern nicht schon alles umgeschifftet wurde
                        If difference > 0 Then

                            For i As Integer = -1 To -xWerteIndex - offset Step -1

                                xWerte(xWerteIndex + offset + i) = xWerte(xWerteIndex + offset + i) - difference


                                If xWerte(xWerteIndex + offset + i) < 0 Then


                                    If msgResult = MsgBoxResult.Yes Then

                                        ' jetzt muss der newValue entsprechend geändert werden 
                                        ' plus, weil xWerte(..) < 0 
                                        newValue = newValue + xWerte(xWerteIndex + offset + i)

                                        ' jetzt muss der newDblValue in das Feld geschrieben werden 
                                        CType(meWS.Cells(targetRow, targetColumn + 2 * i), Excel.Range).Value = newValue
                                        ' die Monatszahl und dann die Summe updaten ... 
                                        CType(meWS.Cells(zeileOFSummaryRole, targetColumn + 2 * i), Excel.Range).Value = 0

                                        difference = -xWerte(xWerteIndex + offset + i)
                                        Exit For
                                    Else

                                        ' zu wenig abgezogen, wird in nächstem Monat abgezogen
                                        Dim zuwenig As Double = -xWerte(xWerteIndex + offset + i)


                                        ' bestimmen der neuen Differenz 
                                        'ur:4.10..2017: difference = newValue - CDbl(visboZustaende.oldValue)
                                        difference = zuwenig

                                        ' ur: 4.10.2017: hier muss die Verteilung von  "difference"  stattfinden

                                        xWerte(xWerteIndex + offset + i) = 0
                                    End If

                                Else

                                    difference = 0

                                End If


                                ' die Monatszahl und dann die Summe updaten ... 
                                '' ''Dim testdbl As Double = xWerte(xWerteIndex + offset + i)
                                '' ''Call MsgBox(" testdbl = " & testdbl.ToString)
                                CType(meWS.Cells(zeileOFSummaryRole, targetColumn + 2 * i), Excel.Range).Value = xWerte(xWerteIndex + offset + i)

                            Next i

                        End If


                    End If

                    ' das wird nachher über updateSummen gemacht 
                    'tmpSum = CDbl(CType(meWS.Cells(zeileOFSummaryRole, columnRC + 1), Excel.Range).Value)
                    'tmpSum = tmpSum - System.Math.Min(alterWert, difference)
                    'CType(meWS.Cells(zeileOFSummaryRole, columnRC + 1), Excel.Range).Value = tmpSum

                    ' nur wenn die Differenz auch ungleich Null ist, muss geändert werden 
                    If difference <> 0 Then
                        summenChanged = True
                    End If

                Else
                    Call MsgBox("Fehler in Übernahme Daten-Wert ...")
                End If

            End If
        End If
    End Sub

    ''' <summary>
    ''' sorgt dafür, dass bei der ParentRole die Summe entsprechend abgeändert und dann verteilt wird 
    ''' </summary>
    ''' <param name="targetRow"></param>
    ''' <param name="targetColumn"></param>
    ''' <param name="newSumValue"></param>
    ''' <param name="hproj"></param>
    ''' <param name="cPhase"></param>
    ''' <param name="roleName"></param>
    ''' <param name="difference"></param>
    ''' <remarks></remarks>
    Private Sub autoReduceRowOfParentRole(ByVal targetRow As Integer, ByVal targetColumn As Integer, ByRef newSumValue As Double, ByVal difference As Double,
                                             ByVal hproj As clsProjekt, ByVal cPhase As clsPhase, ByVal roleName As String)

        Dim pName As String = hproj.name
        Dim phaseNameID As String = cPhase.nameID

        Dim zeileOFSummaryRole As Integer = findeSammelRollenZeile(pName, phaseNameID, roleName)
        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)


        If zeileOFSummaryRole >= 2 And zeileOFSummaryRole <= visboZustaende.meMaxZeile Then

            Dim formerEE As Boolean = appInstance.EnableEvents
            appInstance.EnableEvents = False

            Dim parentSumme As Double = CDbl(CType(meWS.Cells(zeileOFSummaryRole, targetColumn), Excel.Range).Value)
            Dim parentRoleName As String = CStr(meWS.Cells(zeileOFSummaryRole, columnRC).value)
            Dim parentPhaseName As String = CStr(meWS.Cells(zeileOFSummaryRole, 4).value)
            Dim parentPhaseNameID As String = calcHryElemKey(parentPhaseName, False)
            Dim parentComment As Excel.Comment = CType(meWS.Cells(zeileOFSummaryRole, 4), Excel.Range).Comment
            Dim xWerte() As Double

            If Not IsNothing(parentComment) Then
                phaseNameID = parentComment.Text
            End If

            Dim cParentPhase As clsPhase
            If parentPhaseNameID = phaseNameID Then
                cParentPhase = cPhase
            Else
                cParentPhase = hproj.getPhaseByID(parentPhaseNameID)
            End If


            ' jetzt muss in der Sammel-Rolle aktualisiert werden 
            Dim parentRole As clsRolle = Nothing
            Try
                parentRole = cParentPhase.getRole(parentRoleName)
            Catch ex As Exception

            End Try


            If IsNothing(parentRole) Then
                ' nichts tun 
            Else
                ' der Monatswert muss in der parentRole geändert werden 
                xWerte = parentRole.Xwerte

                If parentSumme >= difference Then
                    parentSumme = parentSumme - difference
                Else
                    Dim korrektur As Double = difference - parentSumme
                    newSumValue = newSumValue - korrektur
                    CType(meWS.Cells(targetRow, targetColumn), Excel.Range).Value = newSumValue
                    difference = parentSumme
                    parentSumme = 0
                End If

                ' neuen Wert im Sheet eintragen 
                CType(meWS.Cells(zeileOFSummaryRole, targetColumn), Excel.Range).Value = parentSumme
                CType(meWS.Cells(zeileOFSummaryRole, targetColumn), Excel.Range).NumberFormat = Format("######0.0  ")
                ' jetzt die Rolle aktualisieren 
                Dim parentPhStart As Integer = hproj.Start + cParentPhase.relStart - 1
                Dim parentPhEnde As Integer = hproj.Start + cParentPhase.relEnde - 1

                Dim ixZeitraum As Integer
                Dim ix As Integer
                Dim breite As Integer
                Call awinIntersectZeitraum(parentPhStart, parentPhEnde, ixZeitraum, ix, breite)

                Dim vSum As Double()
                ReDim vSum(0)
                vSum(0) = parentSumme

                Dim xStartDate As Date
                Dim xEndDate As Date

                If ix = 0 Then
                    xStartDate = cParentPhase.getStartDate
                Else
                    xStartDate = cParentPhase.getStartDate.AddDays(-1 * (cParentPhase.getStartDate.Day - 1)).AddMonths(ix)
                End If

                xEndDate = xStartDate.AddDays(-1 * (xStartDate.Day - 1)).AddMonths(breite).AddDays(-1)

                If DateDiff(DateInterval.Day, cParentPhase.getEndDate, xEndDate) > 0 Then
                    xEndDate = cParentPhase.getEndDate
                End If

                Dim xValues() As Double = cParentPhase.berechneBedarfeNew(xStartDate,
                                                                    xEndDate, vSum, 1)

                If parentRole.Xwerte.Length <> xValues.Length Then
                    For lx As Integer = 0 To breite - 1
                        parentRole.Xwerte(lx + ix) = xValues(lx)
                    Next
                Else
                    For i As Integer = 0 To parentRole.Xwerte.Length - 1
                        parentRole.Xwerte(i) = xValues(i)
                    Next
                End If

                ' in der Zeile aktualisieren
                Call aktualisiereRoleCostInSheet(zeileOFSummaryRole, parentRoleName, True, visboZustaende.meColSD, showRangeLeft, showRangeRight,
                                                 parentPhStart, parentPhEnde, xValues)

            End If

            appInstance.EnableEvents = formerEE

        End If
    End Sub
    Private Sub Tabelle2_Deactivate() Handles Me.Deactivate
        ' Achtung: durch das Wechseln der Windows werden auch die ActiveSheets gewechselt; allerdings werden in diesem Fall dann die 
        ' Deactivate Events nicht aufgerufen. Deswegen sollte diese Aktionen alle in separaten Methoden sein  ... 
        ' das ProjInfo Formular löschen, sofern es angezeigt wird 

        ' tk, 3.4.18 wird ohnehin nicht mehr aufgerufen ....
        ' wird jetzt in backtoProjectBoard , performDeactivateActionsFor.. gemacht 
        ''If Not IsNothing(formProjectInfo1) Then
        ''    formProjectInfo1.Close()
        ''End If

        ''Dim meWS As Excel.Worksheet = _
        ''    CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
        ''    .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

        ''appInstance.EnableEvents = False

        '' jetzt den Schutz aufheben , falls einer definiert ist 
        ''If meWS.ProtectContents Then
        ''    meWS.Unprotect(Password:="x")
        ''End If

        ''Try

        ''    ' jetzt die Spalten Werte merken 
        ''    Try
        ''        massColFontValues(0, 0) = CDbl(CType(meWS.Cells(2, 2), Excel.Range).Font.Size)
        ''        For ik As Integer = 1 To 5
        ''            massColFontValues(0, ik) = CDbl(CType(meWS.Columns(ik), Excel.Range).ColumnWidth)
        ''        Next
        ''    Catch ex As Exception

        ''    End Try


        ''    ' jetzt die Autofilter de-aktivieren ... 
        ''    If CType(meWS, Excel.Worksheet).AutoFilterMode = True Then
        ''        CType(meWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
        ''    End If

        ''    ' jetzt alles löschen 
        ''    Try
        ''        meWS.UsedRange.Clear()
        ''    Catch ex As Exception

        ''    End Try

        ''Catch ex As Exception
        ''    Call MsgBox("Fehler beim Filter zurücksetzen " & vbLf & ex.Message)
        ''End Try

        ''appInstance.EnableEvents = True

    End Sub


    Private Sub Tabelle2_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange

        appInstance.EnableEvents = False

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim pname As String = ""
        Dim rcName As String = ""
        Dim oldRCName As String = ""
        Try
            ' wenn mehr wie eine Zelle selektiert wurde ...
            If Target.Cells.Count > 1 Then
                Target = CType(Target.Cells(1, 1), Excel.Range)
                Target.Select()
            End If

            rcName = CStr(meWS.Cells(Target.Row, columnRC).value)

            If visboZustaende.oldRow > 0 Then
                oldRCName = CStr(meWS.Cells(visboZustaende.oldRow, columnRC).value)
            End If

            ' alte Row merken 
            visboZustaende.oldRow = Target.Row

            If awinSettings.meEnableSorting Then
                ' es können auch nicht zugelassene Zellen selektiert worden sein 
                If Target.Cells.Count = 1 Then

                    If isValidSelection(Target) Then
                        oldColumn = Target.Column
                        oldRow = Target.Row
                        If Not IsNothing(Target.Value) Then
                            visboZustaende.oldValue = CStr(Target.Value)
                        Else
                            visboZustaende.oldValue = ""
                        End If
                    Else
                        CType(appInstance.ActiveSheet.Cells(oldRow, oldColumn), Excel.Range).Select()
                    End If


                Else
                    If isValidSelection(CType(Target.Cells(1, 1), Excel.Range)) Then
                        oldColumn = Target.Column
                        oldRow = Target.Row
                        If Not IsNothing(CType(Target.Cells(1, 1), Excel.Range).Value) Then
                            visboZustaende.oldValue = CStr(CType(Target.Cells(1, 1), Excel.Range).Value)
                        Else
                            visboZustaende.oldValue = ""
                        End If
                        CType(Target.Cells(1, 1), Excel.Range).Select()
                    Else
                        CType(appInstance.ActiveSheet.Cells(oldRow, oldColumn), Excel.Range).Select()
                    End If
                End If

            Else
                ' es können nur zugelassene Zellen selektiert worden sein ...
                oldColumn = Target.Column
                oldRow = Target.Row

                If Not IsNothing(CType(Target.Cells(1, 1), Excel.Range).Value) Then
                    visboZustaende.oldValue = CStr(CType(Target.Cells(1, 1), Excel.Range).Value)
                Else
                    visboZustaende.oldValue = ""
                End If

                If Target.Column = columnRC Then
                    'Call MsgBox("RoleCost")
                Else
                    'Call MsgBox("Data")
                End If

            End If
        Catch ex As Exception
            Call MsgBox("Fehler bei Selection Change, Massen-Edit" & vbLf & ex.Message)
            appInstance.EnableEvents = True
        End Try

        ' in oldRow muss jetzt der entsprechende Projekt-Name ausgelsen werden .. 
        ' folgende Bedingung muss gesichert sein: alle Projekte, die in MassEdit aufgeführt sind , 
        ' sind sowohl in Showprojekte als auch in dbCacheProjekte
        Dim pNameChanged As Boolean = False

        With visboZustaende
            pname = CStr(CType(appInstance.ActiveSheet.Cells(Target.Row, visboZustaende.meColpName), Excel.Range).Value)
            If IsNothing(.lastProject) Then
                ' es wurde bisher kein lastProject geladen 
                If ShowProjekte.contains(pname) Then
                    .lastProject = ShowProjekte.getProject(pname)
                    .lastProjectDB = dbCacheProjekte.getProject(calcProjektKey(pname, .lastProject.variantName))
                    pNameChanged = True
                End If
            ElseIf pname <> .lastProject.name Then
                ' muss neu geholt werden 
                If ShowProjekte.contains(pname) Then
                    .lastProject = ShowProjekte.getProject(pname)
                    .lastProjectDB = dbCacheProjekte.getProject(calcProjektKey(pname, .lastProject.variantName))
                    pNameChanged = True
                End If
            End If

            ' wenn pNameChanged und das Info-Fenster angezeigt wird, dann aktualisieren 
            Dim alreadyDone As Boolean = False
            If pNameChanged Then

                selectedProjekte.Clear(False)
                selectedProjekte.Add(.lastProject, False)

                Call aktualisiereCharts(.lastProject, True)

                If Not IsNothing(rcName) Then

                    If rcName <> "" Then
                        Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcName)
                        alreadyDone = True
                    End If
                End If


                If Not IsNothing(formProjectInfo1) Then
                    Call updateProjectInfo1(.lastProject, .lastProjectDB)
                    ' hier wird dann ggf noch das Projekt-/RCNAme/aktuelle Version vs DB-Version Chart aktualisiert  
                End If
            End If


            ' hier wird jetzt ggf das Role/Cost Portfolio Chart aktualisiert ..
            If Not IsNothing(rcName) Then
                If oldRCName <> rcName Then
                    If rcName <> "" And Not alreadyDone Then
                        selectedProjekte.Clear(False)
                        selectedProjekte.Add(.lastProject, False)
                        Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcName)
                    End If
                End If
            End If

        End With



        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' prüft, ob neuer und alter Wert derselben Kategorie angehören; es darf nur von Kostenart zu Kostenart und von Rolle zu Rolle gewechselt werden 
    ''' </summary>
    ''' <param name="newValue"></param>
    ''' <param name="oldValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isValidRCChange(ByVal newValue As String, ByVal oldValue As String) As Boolean

        Dim tmpValue As Boolean = False

        'If RoleDefinitions.containsName(newValue) Then
        '    If RoleDefinitions.containsName(oldValue) Or oldValue = "" Then
        '        tmpValue = True
        '    End If
        'ElseIf CostDefinitions.containsName(newValue) Then
        '    If CostDefinitions.containsName(oldValue) Or oldValue = "" Then
        '        tmpValue = True
        '    End If
        'End If

        If RoleDefinitions.containsName(newValue) Or CostDefinitions.containsName(newValue) Then
            tmpValue = True
        End If

        isValidRCChange = tmpValue

    End Function


    ''' <summary>
    ''' prüft, ob eine gültige Zelle selektiert wurde ... 
    ''' gültig ist eine Zelle, wenn sie entweder in der RoleCost Spalte ist oder in einer Datenspalte 
    ''' und ausserdem die Zeilennummer zwischen 2 und maxzeilen liegt 
    ''' und ausserdem das Projekt nicht geschützt ist ... 
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isValidSelection(ByVal rng As Excel.Range) As Boolean

        Dim result As Boolean = False

        Try
            If rng.Cells.Count > 1 Then
                result = False
            Else
                ' wenn es sich um ein geschütztes Projekt handelt, dann ist Spalte 2 = FarbeProtected, also ungleich dem 
                Dim chckCell As Excel.Range = CType(appInstance.ActiveSheet.Cells(rng.Row, visboZustaende.meColpName), Excel.Range)

                If CInt(chckCell.Interior.ColorIndex) <> XlColorIndex.xlColorIndexNone Then
                    result = False
                Else
                    If rng.Row >= 2 And rng.Row <= visboZustaende.meMaxZeile Then
                        If rng.Column = columnRC Or (rng.Column = columnRC + 1 And awinSettings.allowSumEditing) Then
                            result = True

                        ElseIf rng.Column >= columnStartData And rng.Column <= columnEndData Then
                            Dim diff As Integer = rng.Column - columnStartData
                            Dim rest As Integer
                            Dim tmpValue As Integer = System.Math.DivRem(diff, 2, rest)

                            If rest = 0 Then
                                If rng.Interior.ColorIndex = XlColorIndex.xlColorIndexNone Then
                                    result = False
                                Else
                                    result = True
                                End If
                            Else
                                result = False
                            End If
                        Else
                            result = False
                        End If
                    Else
                        result = False
                    End If
                End If

            End If
        Catch ex As Exception

        End Try


        isValidSelection = result

    End Function

    ''' <summary>
    ''' aktualisiert die Werte in der angegebenen Zeile mit den Daten der Rolle
    ''' der Auslastungs-Array wird in dieser Methode aktualisiert   
    ''' </summary>
    ''' <param name="zeile"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="phStart">ist pStart+relstart-1</param>
    ''' <param name="phEnd">ist pStart+relende -1</param>
    ''' <param name="xWerte"></param>
    ''' <remarks></remarks>
    Private Sub aktualisiereRoleCostInSheet(ByVal zeile As Integer, ByVal rcName As String, ByVal isRole As Boolean,
                                      ByVal startSpalteDaten As Integer,
                                      ByVal von As Integer, ByVal bis As Integer,
                                      ByVal phStart As Integer, ByVal phEnd As Integer,
                                      ByVal xWerte() As Double)
        Dim schnittmenge() As Double
        Dim zeilenWerte() As Double
        Dim zeilensumme As Double
        Dim editRange As Excel.Range



        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' sicherstellen, dass die Länge von xWerte = phStart-phEnd +1 ist
        ' sonst funktioniert die Zuweisung weiter unten nicht 
        If phStart < von Then
            phStart = von
        End If
        If phEnd > bis Then
            phEnd = bis
        End If

        ' wird nur benötigt im Falle isRole ... 
        Dim roleCollection As New Collection
        Dim roleUID As Integer
        Dim auslastungsArray(,) As Double = Nothing

        If isRole Then
            roleUID = RoleDefinitions.getRoledef(rcName).UID
            roleCollection.Add(rcName)
            If awinSettings.meExtendedColumnsView Then
                auslastungsArray = visboZustaende.getUpDatedAuslastungsArray(roleCollection, von, bis, awinSettings.mePrzAuslastung)
            End If

        End If


        Dim ixZeitraum As Integer
        Dim ix As Integer
        Dim breite As Integer
        Call awinIntersectZeitraum(phStart, phEnd, ixZeitraum, ix, breite)

        schnittmenge = calcArrayIntersection(von, bis, phStart, phEnd, xWerte)
        zeilensumme = schnittmenge.Sum

        ReDim zeilenWerte(2 * (bis - von + 1) - 1)

        With CType(appInstance.ActiveSheet, Excel.Worksheet)
            If isRole Then
                If awinSettings.meExtendedColumnsView Then
                    If awinSettings.mePrzAuslastung Then
                        CType(.Cells(zeile, 7), Excel.Range).Value = auslastungsArray(roleUID - 1, 0).ToString("0%")
                    Else
                        CType(.Cells(zeile, 7), Excel.Range).Value = auslastungsArray(roleUID - 1, 0).ToString("#,##0")
                    End If
                Else
                    CType(.Cells(zeile, 7), Excel.Range).Value = ""
                End If

            Else
                CType(.Cells(zeile, 7), Excel.Range).Value = ""
            End If

            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
        End With

        ' zusammenmischen von Schnittmenge und Prozentual-Werte 
        For mis As Integer = 0 To bis - von
            zeilenWerte(2 * mis) = schnittmenge(mis)
            ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung
            If isRole And awinSettings.meExtendedColumnsView Then
                zeilenWerte(2 * mis + 1) = auslastungsArray(roleUID - 1, mis + 1)
            Else
                zeilenWerte(2 * mis + 1) = 0
            End If

        Next

        If awinSettings.meExtendedColumnsView Then
            editRange.Value = zeilenWerte
        Else
            For mis As Integer = 0 To bis - von
                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    CType(.Cells(zeile, startSpalteDaten + 2 * mis), Excel.Range).Value = zeilenWerte(2 * mis)
                    CType(.Cells(zeile, startSpalteDaten + 2 * mis + 1), Excel.Range).Value = zeilenWerte(2 * mis + 1)
                End With
            Next
        End If


        ' jetzt werden die Zellenwerte noch gelöscht , die nicht zur Phase gehören ...  
        With CType(appInstance.ActiveSheet, Excel.Worksheet)
            For l As Integer = 0 To bis - von

                If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                    If isRole Then
                        ' nichts tun 
                    Else
                        ' Auslastung auf Blank setzen 
                        If awinSettings.meExtendedColumnsView Then
                            CType(.Cells(zeile, 2 * l + startSpalteDaten + 1), Excel.Range).Value = ""
                        End If

                    End If
                Else
                    ' diese Werte löschen, sie gehören nicht zum Zeitraum der Phase  
                    CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""

                    If awinSettings.meExtendedColumnsView Then
                        CType(.Cells(zeile, 2 * l + startSpalteDaten + 1), Excel.Range).Value = ""
                    End If

                End If

            Next
        End With

        appInstance.EnableEvents = formerEE

    End Sub


    ''' <summary>
    ''' prüft den Input, setzt, wenn ok, den neuen Wert und die Differenz zum alten Wert 
    ''' </summary>
    ''' <param name="target"></param>
    ''' <param name="newDblValue"></param>
    ''' <param name="difference"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function inputIsAcknowledged(ByVal target As Excel.Range,
                                                ByRef newDblValue As Double,
                                                ByRef difference As Double) As Boolean

        Dim ok As Boolean = False
        ' Bestimmen des Wertes 
        newDblValue = 0.0
        Try
            If IsNothing(target.Cells(1, 1).value) Then
                newDblValue = 0.0
            ElseIf IsNumeric(target.Cells(1, 1).value) Then
                newDblValue = CDbl(target.Cells(1, 1).value)
                If newDblValue >= 0 Then
                    ok = True
                Else
                    newDblValue = 0
                End If
            Else
                newDblValue = 0.0
            End If
        Catch ex As Exception
            newDblValue = 0.0
        End Try

        Try
            If ok Then
                If IsNothing(visboZustaende.oldValue) Then
                    difference = newDblValue
                    visboZustaende.oldValue = "0"
                ElseIf visboZustaende.oldValue = "" Then
                    difference = newDblValue
                    visboZustaende.oldValue = "0"
                Else
                    difference = newDblValue - CDbl(visboZustaende.oldValue)
                End If
            End If

        Catch ex As Exception
            difference = newDblValue
            visboZustaende.oldValue = "0"
        End Try

        inputIsAcknowledged = ok

    End Function

End Class
