
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel


Public Class Tabelle2

    Private columnStartData As Integer = 7
    Private columnEndData As Integer = 18
    Private columnRC As Integer = 5
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private columnName As Integer = 2


    Private Sub Tabelle2_ActivateEvent() Handles Me.ActivateEvent


        Application.DisplayFormulaBar = False

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
                .meMaxZeile = CType(meWS, Excel.Worksheet).UsedRange.Rows.Count
                .meColRC = CType(meWS.Range("RoleCost"), Excel.Range).Column
                .meColSD = CType(meWS.Range("StartData"), Excel.Range).Column
                .meColED = CType(meWS.Range("EndData"), Excel.Range).Column
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

            ' jetzt die Autofilter aktivieren ... 
            If Not CType(meWS, Excel.Worksheet).AutoFilterMode = True Then

                CType(meWS, Excel.Worksheet).Rows(1).AutoFilter()

            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Filtersetzen und Speichern" & vbLf & ex.Message)
        End Try

        Try
            If awinSettings.meEnableSorting Then

                With CType(meWS, Excel.Worksheet)
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

        Application.EnableEvents = formerEE

        ' tk 4.1.20 das wird hier nicht mehr gebracuht, weil Spalte 1 immer selektierbar ist ... 
        ' einen Select machen - nachdem Event Behandlung wieder true ist, dann werden project und lastprojectDB gesetzt ...

        CType(CType(meWS, Excel.Worksheet).Cells(2, 1), Excel.Range).Select()


        ' jetzt die Gridline zeigen
        With appInstance.ActiveWindow
            If massColFontValues(0, 0) <> 0 Then
                .Zoom = massColFontValues(0, 0)
            End If

            .DisplayGridlines = True
            .GridlineColor = Excel.XlRgbColor.rgbBlack
        End With

        ' den alten Wert merken
        If Not IsNothing(appInstance.ActiveCell) Then
            visboZustaende.oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If


        If Application.ScreenUpdating = False Then
            Application.ScreenUpdating = True
        End If

        Dim a As Boolean = appInstance.ScreenUpdating


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

        ' prüfen, ob sich die selektierte Zelle in der Role-/Cost Spalte befindet 
        If Target.Column = columnRC Or Target.Column = columnRC + 1 Then

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
                    ' 
                    Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)
                    Dim skillName As String = CStr(meWS.Cells(zeile, columnRC + 1).value)

                    Dim rcNameID As String = getRCNameIDfromExcelRange(CType(meWS.Range(Cells(zeile, columnRC), Cells(zeile, columnRC + 1)), Excel.Range))
                    Dim phaseNameID As String = getPhaseNameIDfromExcelCell(CType(meWS.Cells(zeile, columnRC - 1), Excel.Range))


                    Dim hproj As clsProjekt = Nothing
                    If Not IsNothing(pName) Then
                        If pName <> "" Then
                            hproj = ShowProjekte.getProject(pName)
                        End If
                    End If

                    ' es handelt sich um eine Rollen- oder Kosten-Änderung ...
                    ' Jetzt muss ein Formular mit den Rollen und Kosten im TreeView angezeigt werden
                    If IsNothing(pName) Then
                        pName = ""
                    End If
                    If IsNothing(vName) Then
                        vName = ""
                    End If
                    If IsNothing(phaseName) Then
                        phaseName = ""
                    End If
                    If IsNothing(rcName) Then
                        rcName = ""
                        rcNameID = ""
                    End If
                    If IsNothing(phaseNameID) Then
                        phaseNameID = ""
                    End If

                    If IsNothing(skillName) Then
                        skillName = ""
                    End If



                    frmMERoleCost.pName = pName
                    frmMERoleCost.vName = vName
                    frmMERoleCost.phaseName = phaseName
                    frmMERoleCost.rcName = rcName
                    frmMERoleCost.rcNameID = rcNameID
                    frmMERoleCost.phaseNameID = phaseNameID
                    frmMERoleCost.skillName = skillName

                    If Target.Column = columnRC Then
                        frmMERoleCost.showSkillsOnly = False
                    Else
                        frmMERoleCost.showSkillsOnly = True
                    End If

                    frmMERoleCost.hproj = hproj

                    returnValue = frmMERoleCost.ShowDialog()

                    If returnValue = DialogResult.OK Then


                        For Each roleSkillItem As String In frmMERoleCost.rolesToAdd
                            If frmMERoleCost.showSkillsOnly Then
                                If rcName = "" Then

                                    Try
                                        Dim tmpID As Integer = -1
                                        rcName = RoleDefinitions.getContainingRoleOfSkillMembers(RoleDefinitions.getRoleDefByIDKennung(roleSkillItem, tmpID).UID).name
                                    Catch ex As Exception

                                    End Try

                                End If
                                Call meRCZeileEinfuegen(zeile, rcName, roleSkillItem, True)
                            Else
                                Call meRCZeileEinfuegen(zeile, roleSkillItem, skillName, True)
                            End If


                            zeile = visboZustaende.oldRow
                        Next

                        For Each costNameIDitem As String In frmMERoleCost.costsToAdd
                            Call meRCZeileEinfuegen(zeile, costNameIDitem, "", False)

                            zeile = visboZustaende.oldRow
                        Next



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

        ' ColumnRC + 1 steht jetzt immer der Skill 

        Dim currentCell As Excel.Range = Target

        Try
            Dim auslastungChanged As Boolean = False
            Dim summenChanged As Boolean = False
            ' muss extra überwacht werden, um das ProjectInfo1 Fenster auch immer zu aktualisieren
            Dim kostenChanged As Boolean = False


            Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
            Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

            If Target.Cells.Count = 1 Or Target.Rows.Count = 1 Then


                Dim roleCostNames As New Collection

                Dim isRole As Boolean = False
                Dim isCost As Boolean = False

                Dim zeile As Integer = Target.Row
                Dim pName As String = CStr(meWS.Cells(zeile, visboZustaende.meColpName).value)
                Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
                Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
                Dim rcName As String = ""
                Dim rcNameID As String = ""

                If Not IsNothing(meWS.Cells(zeile, columnRC).value) Then
                    rcName = CStr(meWS.Cells(zeile, columnRC).value).Trim
                    If rcName <> "" Then
                        isRole = RoleDefinitions.containsName(rcName)
                        If Not isRole Then
                            isCost = CostDefinitions.containsName(rcName)
                        End If
                    End If
                End If

                Dim skillName As String = ""
                Dim oldSkillID As Integer = -1

                If Not IsNothing(meWS.Cells(zeile, columnRC + 1).value) Then
                    skillName = CStr(meWS.Cells(zeile, columnRC + 1).value).Trim
                    If skillName.Length > 0 Then
                        If RoleDefinitions.containsName(skillName) Then
                            oldSkillID = RoleDefinitions.getRoledef(skillName).UID
                        End If
                    End If

                End If

                If isRole Then
                    rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                End If


                Dim phaseNameID As String = getPhaseNameIDfromExcelCell(CType(meWS.Cells(zeile, columnRC - 1), Excel.Range))

                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                Dim cphase As clsPhase = Nothing


                If Target.Columns.Count = 1 Then

                    If Not IsNothing(hproj) Then
                        cphase = hproj.getPhaseByID(phaseNameID)
                        If Not IsNothing(cphase) Then

                            If Target.Column = columnRC Then
                                ' es handelt sich um eine Rollen- oder Kosten-Änderung ...

                                ' steht jetzt in rcNAme 
                                'newRCName = CStr(Target.Cells(1, 1).value)

                                If IsNothing(rcName) Then
                                    If Not IsNothing(visboZustaende.oldValue) Then
                                        If visboZustaende.oldValue <> "" Then
                                            Dim errMsg As String = "um Rolle /Kostenart zu löschen bitte entsprechenden Menupunkt nutzen ... "
                                            If awinSettings.englishLanguage Then
                                                errMsg = "to delete a role or cost, please use the according menu-item ..."
                                            End If
                                            Call MsgBox(errMsg)
                                            Target.Cells(1, 1).value = visboZustaende.oldValue
                                        End If
                                    End If

                                ElseIf rcName.Trim = "" Then

                                    If visboZustaende.oldValue <> "" Then
                                        Dim errMsg As String = "um Rolle /Kostenart zu löschen bitte entsprechenden Menupunkt nutzen ... "
                                        If awinSettings.englishLanguage Then
                                            errMsg = "to delete a role or cost, please use the according menu-item ..."
                                        End If
                                        Call MsgBox(errMsg)
                                        Target.Cells(1, 1).value = visboZustaende.oldValue
                                    End If


                                ElseIf isValidRCChange(rcName, visboZustaende.oldValue, skillName, False) Then
                                    ' es ist eine gültige Änderung, das heisst es wurde eine Rolle in eine andere gewechselt , oder 
                                    ' eine Kostenart in eine andere; Kategorie-übergreifende Wechsel sind nicht erlaubt 

                                    ' jetzt muss noch geprüft werden, ob auch keine Duplikate vorkommen: zu einem Projekt dürfen z.Bsp keine 
                                    ' 2 Zeilen existieren mit jeweils der gleichen Rolle oder Kostenart ...
                                    If noDuplicatesInSheet(pName, phaseNameID, rcNameID, zeile) Then

                                        Dim rcIndentLevel As Integer = 0
                                        If isRole Then
                                            ' es handelt sich um eine Rollen-Änderung

                                            Dim skillID As Integer = -1
                                            Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(rcNameID, skillID)

                                            ' jetzt den Indentlevel der Rolle vestimmen bestimmen 
                                            rcIndentLevel = RoleDefinitions.getRoleIndent(rcNameID)
                                            currentCell.IndentLevel = rcIndentLevel

                                            Dim newRoleID As Integer = tmpRole.UID
                                            If visboZustaende.oldValue.Trim.Length > 0 And visboZustaende.oldValue.Trim <> rcName.Trim Then
                                                ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                                Try
                                                    auslastungChanged = True
                                                    Dim cRole As clsRolle = cphase.getRole(visboZustaende.oldValue, oldSkillID)
                                                    If IsNothing(cRole) Then
                                                    Else
                                                        'hproj.rcLists.removeRP(cRole.uid, cphase.nameID, skillID, False)
                                                        cRole.uid = newRoleID
                                                        'hproj.rcLists.addRP(newRoleID, cphase.nameID, skillID)
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



                                            ' tk 23.8.2020 das wird doch nicht mehr benötigt ... 
                                            ' jetzt für die Zelle die Validation neu bestimmen, dazu muss aber der Blattschutz aufgehoben sein ...  

                                            'If Not awinSettings.meEnableSorting Then
                                            '    ' es muss der Blattschutz aufgehoben werden, nachher wieder mit diesen Einstellungen aktiviert werden ...
                                            '    With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                            '        .Unprotect(Password:="x")
                                            '    End With
                                            'End If


                                            'If Not awinSettings.meEnableSorting Then
                                            '    ' es muss der Blattschutz aufgehoben werden, nachher wieder mit diesen Einstellungen aktiviert werden ...
                                            '    With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                            '        .Protect(Password:="x", UserInterfaceOnly:=True,
                                            '                 AllowFormattingCells:=True,
                                            '                 AllowFormattingColumns:=True,
                                            '                 AllowInsertingColumns:=False,
                                            '                 AllowInsertingRows:=True,
                                            '                 AllowDeletingColumns:=False,
                                            '                 AllowDeletingRows:=True,
                                            '                 AllowSorting:=True,
                                            '                 AllowFiltering:=True)
                                            '        .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
                                            '        .EnableAutoFilter = True
                                            '    End With
                                            'End If

                                            'If awinSettings.meAutoReduce Then
                                            '    ' jetzt die Rollen bestimmen, die neu berechnet werden müssen ... 
                                            '    roleCostNames = RoleDefinitions.getSummaryRoles(rcName)
                                            '    If Not roleCostNames.Contains(rcName) Then
                                            '        roleCostNames.Add(rcName, rcName)
                                            '    End If

                                            '    If visboZustaende.oldValue.Length > 0 Then
                                            '        If Not roleCostNames.Contains(visboZustaende.oldValue) Then
                                            '            roleCostNames.Add(visboZustaende.oldValue, visboZustaende.oldValue)
                                            '        End If

                                            '        Dim tmpSummaryNames As Collection = RoleDefinitions.getSummaryRoles(visboZustaende.oldValue)
                                            '        For sr As Integer = 1 To tmpSummaryNames.Count
                                            '            Dim srName As String = CStr(tmpSummaryNames.Item(sr))
                                            '            If Not roleCostNames.Contains(srName) Then
                                            '                roleCostNames.Add(srName, srName)
                                            '            End If
                                            '        Next
                                            '    End If
                                            'End If

                                            ' Ende Rollen-Behandlung
                                        ElseIf isCost Then
                                            ' es handelt sich um eine Kostenart Änderung 
                                            If visboZustaende.oldValue.Length > 0 And visboZustaende.oldValue.Trim <> rcName.Trim Then
                                                ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                                Dim newCostID As Integer = CostDefinitions.getCostdef(rcName).UID
                                                Dim cCost As clsKostenart = cphase.getCost(visboZustaende.oldValue)
                                                If IsNothing(cCost) Then
                                                Else
                                                    'hproj.rcLists.removeCP(cCost.KostenTyp, cphase.nameID)
                                                    cCost.KostenTyp = newCostID
                                                    'hproj.rcLists.addCP(newCostID, cphase.nameID)
                                                End If
                                                kostenChanged = True
                                            Else
                                                ' es kam eine neue Kostenart hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
                                                ' muss an dieser Stelle noch gar nichts gemacht werden ..
                                            End If
                                        Else
                                            ' falsche/unbekannte Eingabe
                                            Dim errMsg As String = "unbekannte Rolle / Kostenart ..."
                                            If awinSettings.englishLanguage Then
                                                errMsg = "unknown role/cost ..."
                                            End If
                                            Call MsgBox(errMsg)

                                            Target.Cells(1, 1).value = visboZustaende.oldValue
                                        End If


                                    Else
                                        Dim errMsg As String = "keine Doppelbelegung innerhalb einer Projektphase erlaubt ... "
                                        If awinSettings.englishLanguage Then
                                            errMsg = "no duplicates within one phase, please"
                                        End If
                                        Call MsgBox(errMsg)

                                        Target.Cells(1, 1).value = visboZustaende.oldValue

                                        If visboZustaende.oldValue = "" Or IsNothing(visboZustaende.oldValue) Then
                                            ' Zeile löschen mit Doppelbelegung
                                            Call massEditZeileLoeschen("")

                                        ElseIf RoleDefinitions.containsName(visboZustaende.oldValue) Then
                                            Target.ClearComments()

                                        End If


                                    End If

                                    ' jetzt muss in der Spaltenüberschrift noch angegeben werden, ob es sich um T€, PT oder nichts handelt 
                                    Call defineHeaderTitleOfRoleCost(Target.Row)




                                Else
                                    Dim errMsg As String = "unbekannte Rolle / Kostenart ..."
                                    If awinSettings.englishLanguage Then
                                        errMsg = "unknown role/cost ..."
                                    End If
                                    Call MsgBox(errMsg)
                                    Target.Cells(1, 1).value = visboZustaende.oldValue
                                End If

                            ElseIf Target.Column = columnRC + 1 Then
                                ' es handelt sich um eine Skill Änderung


                                If isValidRCChange(skillName, visboZustaende.oldValue, rcName, True) Then

                                    isRole = True

                                    Dim rcNameGenerated As Boolean = False
                                    If rcName = "" Then
                                        rcName = RoleDefinitions.getContainingRoleOfSkillMembers(oldSkillID).name
                                        rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                                        rcNameGenerated = True
                                    End If

                                    If noDuplicatesInSheet(pName, phaseNameID, rcNameID, zeile) Then

                                        If rcNameGenerated Then
                                            ' der automatisch generierte Name muss jetzt eingetragen werden 
                                            Try
                                                Target.Cells(1, 1).offset(0, -1).value = rcName
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        oldSkillID = -1
                                        Dim newSkillID As Integer = -1
                                        If visboZustaende.oldValue <> "" Then
                                            If RoleDefinitions.containsName(visboZustaende.oldValue) Then
                                                oldSkillID = RoleDefinitions.getRoledef(visboZustaende.oldValue).UID
                                            End If
                                        End If

                                        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(rcNameID, newSkillID)

                                        If newSkillID <> oldSkillID Then
                                            ' es handelt sich um einen Wechsel 
                                            Try
                                                auslastungChanged = True
                                                Dim cRole As clsRolle = cphase.getRole(rcName, oldSkillID)
                                                If IsNothing(cRole) Then
                                                Else
                                                    'hproj.rcLists.removeRP(cRole.uid, cphase.nameID, oldSkillID, False)
                                                    cRole.teamID = newSkillID
                                                    'hproj.rcLists.addRP(cRole.uid, cphase.nameID, newSkillID)
                                                End If


                                            Catch ex As Exception
                                                visboZustaende.oldValue = ""
                                                ' in diesem Fall wurde nur von einer noch nicht belegten Rolle auf eine 
                                                ' andere nicht belegte gewechselt 
                                            End Try
                                        End If

                                        ' bestimme den 

                                    Else
                                        Dim errMsg As String = "keine Doppelbelegung innerhalb einer Projektphase erlaubt ... "
                                        If awinSettings.englishLanguage Then
                                            errMsg = "no duplicates within one phase, please"
                                        End If
                                        Call MsgBox(errMsg)

                                        Target.Cells(1, 1).value = visboZustaende.oldValue

                                    End If
                                Else
                                    Dim errMsg As String = ""
                                    If isCost Then
                                        errMsg = "bei Kostenart ist Skill Angabe nicht zugelassen ... "
                                        If awinSettings.englishLanguage Then
                                            errMsg = "costs do not carry skills ... "
                                        End If
                                    Else
                                        errMsg = "Skill existiert in der angegebenen Organisations-Einheit nicht ..."

                                        If awinSettings.englishLanguage Then
                                            errMsg = "skill does not exist in orga-unit ..."
                                        End If
                                    End If

                                    Call MsgBox(errMsg)
                                    Target.Cells(1, 1).value = visboZustaende.oldValue
                                End If


                            ElseIf Target.Column = columnRC + 2 Then
                                ' es handelt sich um eine Summenänderung
                                Dim newDblValue As Double
                                Dim difference As Double
                                Dim teamID As Integer = -1
                                Dim ok As Boolean = False

                                Dim uid As Integer

                                If isRole Then
                                    Dim roleInRow As clsRollenDefinition = Nothing
                                    roleInRow = RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID)
                                    If Not IsNothing(roleInRow) Then
                                        uid = roleInRow.UID
                                        ok = True
                                    End If

                                ElseIf isCost Then
                                    Dim costInRow As clsKostenartDefinition = Nothing
                                    costInRow = CostDefinitions.getCostdef(rcName)
                                    If Not IsNothing(costInRow) Then
                                        uid = costInRow.UID
                                        ok = True
                                    End If

                                Else
                                    Dim errMsg As String = "bitte erst eine Rolle oder Kostenart auswählen ..."
                                    If awinSettings.englishLanguage Then
                                        errMsg = "please, first choose a role or cost name ..."
                                    End If
                                    Call MsgBox(errMsg)
                                    Target.Cells(1, 1).value = visboZustaende.oldValue
                                End If

                                If ok Then

                                    If inputIsAcknowledged(Target, newDblValue, difference) Then

                                        Dim phStart As Integer = hproj.Start + cphase.relStart - 1
                                        Dim phEnde As Integer = hproj.Start + cphase.relEnde - 1

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
                                            xStartDate = cphase.getStartDate
                                        Else
                                            xStartDate = cphase.getStartDate.AddDays(-1 * (cphase.getStartDate.Day - 1)).AddMonths(ix)
                                        End If

                                        xEndDate = xStartDate.AddDays(-1 * (xStartDate.Day - 1)).AddMonths(breite).AddDays(-1)

                                        If DateDiff(DateInterval.Day, cphase.getEndDate, xEndDate) > 0 Then
                                            xEndDate = cphase.getEndDate
                                        End If

                                        Dim xValues() As Double = cphase.berechneBedarfeNew(xStartDate,
                                                                                                    xEndDate, vSum, 1)

                                        If isRole Then

                                            ' erstmal überprüfen, ob awinsettings.autoreduce = true 
                                            Dim parentRoleSum As Double = -1
                                            If awinSettings.meAutoReduce Then
                                                Call autoReduceRowOfParentRole(Target.Row, Target.Column, newDblValue, difference,
                                                                                       hproj, cphase, rcName)

                                                ' durch autoReduce kann der newDblValue verändert sein
                                                vSum(0) = newDblValue
                                                xValues = cphase.berechneBedarfeNew(xStartDate, xEndDate, vSum, 1)

                                            End If

                                            ' jetzt muss die Rolle aktualisiert werden ...
                                            Dim tmpRole As clsRolle = cphase.getRoleByRoleNameID(rcNameID)

                                            If IsNothing(tmpRole) Then
                                                tmpRole = New clsRolle(phEnde - phStart)

                                                With tmpRole
                                                    .uid = uid
                                                    .teamID = teamID
                                                End With
                                                With cphase
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
                                            Dim tmpCost As clsKostenart = cphase.getCost(rcName)

                                            If IsNothing(tmpCost) Then
                                                tmpCost = New clsKostenart(phEnde - phStart)

                                                With tmpCost
                                                    .KostenTyp = uid
                                                End With
                                                With cphase
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

                                    Else
                                        ' nichts tun 
                                    End If

                                End If



                            ElseIf Target.Column > columnRC + 2 Then



                                ' es handelt sich um eine Datenänderung
                                'Dim newDblValue As Double
                                'Dim difference As Double

                                ' zu welcher / welchen Sammelrollen gehört die ausgewählte Rolle ? 
                                Dim sammelRollenName As String = ""
                                Dim zeileSammelRolle As Integer = 0


                                If isRole Or isCost Then
                                    ' hier ist etwas gültiges vorhanden .. es kann also weitergemacht werden 

                                    Call updateDataValuesInProject(Target, isRole, rcName, rcNameID, pName, phaseNameID,
                                                                auslastungChanged, summenChanged, kostenChanged, roleCostNames)


                                Else
                                    Dim errMsg As String = "bitte erst eine Rolle oder Kostenart auswählen ..."
                                    If awinSettings.englishLanguage Then
                                        errMsg = "please, first choose a role or cost name ..."
                                    End If
                                    Call MsgBox(errMsg)
                                    Target.Cells(1, 1).value = visboZustaende.oldValue
                                End If

                            Else
                                ' es wurde die Business Unit selektiert ..
                                Target.Cells(1, 1).value = visboZustaende.oldValue
                            End If

                        End If

                    End If

                ElseIf Target.Columns.Count > 1 Then

                    If isRole Or isCost Then
                        Call updateDataValuesInProject(Target, isRole, rcName, rcNameID, pName, phaseNameID,
                                                                auslastungChanged, summenChanged, kostenChanged, roleCostNames)
                    End If

                End If




                'If auslastungChanged And awinSettings.meExtendedColumnsView Then
                '    'Call updateMassEditAuslastungsValues(showRangeLeft, showRangeRight, roleCostNames)
                'End If

                ' das Folgende ist eigentlich eine Test Routine , die normalerweise gar nicht nötig ist 
                ' aber für Testzwecke gut geeignet ist ...

                'Dim testValue1 As Double = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)
                If summenChanged Then

                        If IsNothing(cphase) Then
                            ' wenn in Zweig target.columns.count > 1 gewesen
                            cphase = hproj.getPhaseByID(phaseNameID)
                        End If

                        Call updateMassEditSummenValue(hproj, cphase, showRangeLeft, showRangeRight, rcNameID, isRole, zeile)

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
                                Call updateProjectInfo1(visboZustaende.currentProject, visboZustaende.currentProjectinSession)
                            End If
                        ' tk 18.1.20
                        Call aktualisiereCharts(visboZustaende.currentProject, True, calledFromMassEdit:=True, currentRCName:=rcName)
                        'Call aktualisiereCharts(visboZustaende.currentProject, True, calledFromMassEdit:=True, currentRoleName:=rcName)

                        Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcNameID)
                    End If

                    Catch ex As Exception

                    End Try

                ElseIf Target.Rows.Count > 1 Then

                    'appInstance.Undo()
                    'Call MsgBox("bitte nur eine Zelle selektieren ...")
                    appInstance.Undo()
                'Target.Cells(1, 1).value = visboZustaende.oldValue
            End If


        Catch ex As Exception
            Dim errMsg As String = "Fehler bei Massen-Edit, Ändern : " & vbLf & ex.Message
            If awinSettings.englishLanguage Then
                errMsg = "Error in editing resources / cost: " & vbLf & ex.Message
            End If
            Call MsgBox(errMsg)
        End Try

        appInstance.EnableEvents = True
    End Sub

    ''' <summary>
    ''' aktualisiert 
    ''' </summary>
    ''' <param name="target"></param>
    ''' <param name="isRole"></param>
    ''' <param name="rcName"></param>
    ''' <param name="rcNameID"></param>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="auslastungChanged"></param>
    ''' <param name="summenchanged"></param>
    ''' <param name="kostenchanged"></param>
    ''' <param name="roleCostNames"></param>
    Private Sub updateDataValuesInProject(ByVal target As Excel.Range,
                                        ByVal isRole As Boolean,
                                        ByVal rcName As String,
                                        ByVal rcNameID As String,
                                        ByVal pName As String,
                                        ByVal phaseNameID As String,
                                        ByRef auslastungChanged As Boolean,
                                        ByRef summenchanged As Boolean,
                                        ByRef kostenchanged As Boolean,
                                        ByRef roleCostNames As Collection)

        ' es handelt sich um eine Datenänderung
        Dim newDblValue As Double
        Dim difference As Double

        Dim anzTargetColumns As Integer = target.Columns.Count


        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
        If Not IsNothing(hproj) Then
            Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)

            If Not IsNothing(cphase) Then
                ' hier ist etwas gültiges vorhanden .. es kann also weitergemacht werden 

                Try
                    If IsNothing(target.Cells(1, 1).value) Then
                        newDblValue = 0.0
                    ElseIf IsNumeric(target.Cells(1, 1).value) Then
                        newDblValue = CDbl(target.Cells(1, 1).value)
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

                Dim monthCol As Integer = showRangeLeft + target.Column - columnStartData

                Dim xWerteIndex As Integer = monthCol - getColumnOfDate(cphase.getStartDate)
                Dim xWerte() As Double

                If isRole Then
                    ' es handelt sich um eine gültige Rolle

                    If awinSettings.meAutoReduce Then

                        Call autoReduceCellOfParentRole(target.Row, target.Column, newDblValue,
                                                              hproj, cphase, rcName, xWerteIndex, difference, summenchanged)

                    End If

                    ' es muss einfach die Rolle hinzugefügt bzw. die Werte abgeändert werden 
                    Dim tmpRole As clsRolle = cphase.getRoleByRoleNameID(rcNameID)

                    If IsNothing(tmpRole) Then
                        ' die Rolle muss neu angelegt und der Phase hinzugefügt werden  

                        tmpRole = New clsRolle(cphase.relEnde - cphase.relStart)
                        Dim teamID As Integer = -1
                        tmpRole.uid = RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID).UID
                        tmpRole.teamID = teamID

                        Call cphase.addRole(tmpRole)

                    End If

                    ' der Monatswert muss geändert werden 
                    xWerte = tmpRole.Xwerte

                    For i As Integer = 1 To anzTargetColumns
                        If xWerteIndex >= 0 And xWerteIndex <= xWerte.Length - 1 Then
                            If xWerte(xWerteIndex) <> newDblValue Then
                                xWerte(xWerteIndex) = newDblValue
                                summenchanged = True
                            End If
                        Else
                            ' nichts weiter tun, ausserhalb Werte Bereich
                            Exit For
                        End If
                        xWerteIndex = xWerteIndex + 1
                    Next


                    ' bestimmt zu welchen Rollen die Auslastungs-Werte neu berechnet werden müssen ..
                    If awinSettings.meAutoReduce Then
                        roleCostNames = RoleDefinitions.getSummaryRoles(rcName)
                        If Not roleCostNames.Contains(rcName) Then
                            roleCostNames.Add(rcName, rcName)
                        End If
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

                        kostenchanged = True
                    End If

                    ' der Monatswert muss geändert werden 
                    xWerte = tmpCost.Xwerte

                    For i As Integer = 1 To anzTargetColumns
                        If xWerteIndex >= 0 And xWerteIndex <= xWerte.Length - 1 Then
                            If xWerte(xWerteIndex) <> newDblValue Then
                                xWerte(xWerteIndex) = newDblValue
                                summenchanged = True
                            End If
                        Else
                            ' nichts weiter tun, ausserhalb Werte Bereich
                            Exit For
                        End If
                        xWerteIndex = xWerteIndex + 1
                    Next

                    '
                    If awinSettings.meAutoReduce Then
                        If Not roleCostNames.Contains(rcName) Then
                            roleCostNames.Add(rcName, rcName)
                        End If
                    End If


                End If
            End If

        End If





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
                            CType(meWS.Cells(zeileOFSummaryRole, targetColumn + i), Excel.Range).Value = xWerte(xWerteIndex + offset + i)

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
                                        CType(meWS.Cells(targetRow, targetColumn + i), Excel.Range).Value = newValue
                                        ' die Monatszahl und dann die Summe updaten ... 
                                        CType(meWS.Cells(zeileOFSummaryRole, targetColumn + i), Excel.Range).Value = 0

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
                                CType(meWS.Cells(zeileOFSummaryRole, targetColumn + i), Excel.Range).Value = xWerte(xWerteIndex + offset + i)

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

        appInstance.ActiveWindow.SplitColumn = 0
        appInstance.ActiveWindow.SplitRow = 0
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

    ''' <summary>
    ''' gibt in der Headerzeile an, ob es sich bei den Werten in der Zeile um Personentage oder oder um Tausend Euro handelt 
    ''' </summary>
    ''' <param name="zeile"></param>
    Private Sub defineHeaderTitleOfRoleCost(ByVal zeile As Integer)

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim headerPart As String = "Summe" & vbLf
        Dim pdEinheit As String = "[PT]"

        If awinSettings.englishLanguage Then
            headerPart = "Sum" & vbLf
            pdEinheit = "[PD]"
        End If

        Dim roleCost As String = CStr(CType(meWS.Cells(zeile, visboZustaende.meColRC), Excel.Range).Value)

        If Not IsNothing(roleCost) Then
            If roleCost = "" Then
                CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart

            ElseIf RoleDefinitions.containsName(roleCost) Then
                CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart & pdEinheit

            Else
                CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart & "[T€]"
            End If
        Else
            CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart
        End If



    End Sub

    ''' <summary>
    ''' wird aufgerufen, wenn sich die Zeile ändert ...
    ''' </summary>
    ''' <param name="Target"></param>
    Private Sub Tabelle2_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange

        appInstance.EnableEvents = False



        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)

        If Target.Row <> oldRow Then
            'Call highlightRow(Target.Row, oldRow)

            ' jetzt muss in der Spaltenüberschrift noch angegeben werden, ob es sich um T€, PT oder nichts handelt 
            Call defineHeaderTitleOfRoleCost(Target.Row)

        End If


        Dim pname As String = ""
        Dim rcName As String = ""
        Dim rcNameID As String = ""
        Dim oldRCName As String = ""
        Dim oldRCNameID As String = ""

        Dim changeBecauseRCNameChanged As Boolean = False
        Try
            ' wenn mehr wie eine Zelle selektiert wurde ...
            If Target.Cells.Count > 1 Then
                Target = CType(Target.Cells(1, 1), Excel.Range)
                Target.Select()
            End If

            rcName = CStr(meWS.Cells(Target.Row, columnRC).value)
            rcNameID = getRCNameIDfromExcelRange(CType(meWS.Range(meWS.Cells(Target.Row, columnRC), meWS.Cells(Target.Row, columnRC + 1)), Excel.Range))

            If visboZustaende.oldRow > 0 Then
                oldRCName = CStr(meWS.Cells(visboZustaende.oldRow, columnRC).value)
                oldRCNameID = getRCNameIDfromExcelRange(CType(meWS.Range(meWS.Cells(visboZustaende.oldRow, columnRC), meWS.Cells(visboZustaende.oldRow, columnRC + 1)), Excel.Range))
            End If

            ' das wirkt sich auf das aktualisieren der charts aus 
            'changeBecauseRCNameChanged = rcName <> oldRCName
            changeBecauseRCNameChanged = rcNameID <> oldRCNameID

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

            If IsNothing(.currentProject) Then
                ' es wurde bisher kein lastProject geladen 
                If ShowProjekte.contains(pname) Then
                    .currentProject = ShowProjekte.getProject(pname)
                    .currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pname, .currentProject.variantName))
                    pNameChanged = True
                End If

            ElseIf pname <> .currentProject.name Then
                ' muss neu geholt werden 
                If ShowProjekte.contains(pname) Then
                    .currentProject = ShowProjekte.getProject(pname)
                    .currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pname, .currentProject.variantName))
                    pNameChanged = True
                End If
            End If

            ' wenn pNameChanged und das Info-Fenster angezeigt wird, dann aktualisieren 

            If pNameChanged Or changeBecauseRCNameChanged Then

                Call aktualisiereCharts(.currentProject, True, calledFromMassEdit:=True, currentRCName:=rcName)

                If pNameChanged Then
                    selectedProjekte.Clear(False)
                    selectedProjekte.Add(.currentProject, False)
                End If

                If Not IsNothing(rcNameID) Then

                    If rcNameID <> "" Then
                        Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcNameID)
                    End If
                End If


                If Not IsNothing(formProjectInfo1) Then
                    Call updateProjectInfo1(.currentProject, .currentProjectinSession)
                    ' hier wird dann ggf noch das Projekt-/RCNAme/aktuelle Version vs DB-Version Chart aktualisiert  
                End If


            End If


            ' tk 29.10. alt - nur rcname
            'If Not IsNothing(rcName) Then
            '    If oldRCName <> rcName Then
            '        If rcName <> "" And Not alreadyDone Then
            '            selectedProjekte.Clear(False)
            '            selectedProjekte.Add(.currentProject, False)
            '            Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcName)
            '        End If
            '    End If
            'End If

        End With



        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' prüft, ob der eingegebene Wert zulässig ist ..
    ''' ein Ressourcen-Manager darf nur Werte seiner Abteilung eingeben
    ''' ein Portfolio Manager darf niemanden unterhalb der customerrole.specifics auswählen 
    ''' </summary>
    ''' <param name="newValue"></param>
    ''' <param name="oldValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isValidRCChange(ByVal newValue As String, ByVal oldValue As String, ByVal otherValue As String, ByVal isSkillCheck As Boolean) As Boolean

        Dim rcName As String = ""
        Dim skillName As String = ""
        Dim tmpValue As Boolean = False

        If isSkillCheck Then
            ' es handelt sich um den Skill Check 
            rcName = otherValue
            skillName = newValue


            If rcName <> "" Then
                ' prüfen, ob es diese Skill denn überhaupt für die Rolle gibt 
                ' wenn es sich um eine Kostenart handelt : sillok = false
                tmpValue = RoleDefinitions.roleHasSkill(rcName, skillName)
            Else
                ' anderfalls muss geprüft werden ob es sich um eine gültige Skill handelt ... 
                Dim tmpSkill As clsRollenDefinition = RoleDefinitions.getRoledef(skillName)
                If Not IsNothing(tmpSkill) Then
                    tmpValue = tmpSkill.isSkill
                End If
            End If

        Else
            ' es handelt sich um den RCName Check
            rcName = newValue
            skillName = otherValue
            Dim stillOk As Boolean = True

            If skillName <> "" Then
                ' prüfen, ob es diese Skill denn überhaupt für die Rolle gibt 
                ' wenn es sich um eine Kostenart handelt : sillok = false
                stillOk = RoleDefinitions.roleHasSkill(rcName, skillName)
            End If

            If stillOk Then

                Dim weiterMachen As Boolean = False
                Dim skillID As Integer = -1

                ' erstmal prüfen, ob es sich um einen Ressourcen-Manager oder Portfolio Manager handelt; denn dann können nicht alle Werte eingegeben werden 
                If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                    Dim parentCollection As New Collection

                    'parentCollection.Add(myCustomUserRole.specifics)
                    parentCollection.Add(RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, skillID).name)

                    If RoleDefinitions.hasAnyChildParentRelationsship(newValue, parentCollection) Then
                        weiterMachen = True
                    End If

                ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                    'Dim idArray() As Integer = RoleDefinitions.getIDArray(myCustomUserRole.specifics)
                    Dim idArray() As Integer = myCustomUserRole.getAggregationRoleIDs

                    If Not IsNothing(idArray) Then
                        Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(newValue, "")

                        Dim roleID As Integer = RoleDefinitions.parseRoleNameID(roleNameID, skillID)

                        If Not RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, idArray) Or idArray.Contains(roleID) Then
                            weiterMachen = True
                        End If
                    Else
                        weiterMachen = True
                    End If

                Else
                    weiterMachen = True
                End If

                If weiterMachen Then
                    If oldValue.Trim.Length = 0 Then
                        ' ist erlaubt, wenn der Wert in einer der Definitions vorkommt 
                        tmpValue = RoleDefinitions.containsName(newValue) Or CostDefinitions.containsName(newValue)
                    Else
                        ' es war vorher was drin 
                        If RoleDefinitions.containsName(newValue) Or CostDefinitions.containsName(newValue) Then

                            If RoleDefinitions.containsName(newValue) = RoleDefinitions.containsName(oldValue) Then
                                ' ist erlaubt 
                                tmpValue = True
                            ElseIf CostDefinitions.containsName(newValue) = missingCostDefinitions.containsName(oldValue) Then
                                ' ist erlaubt
                                tmpValue = True
                            Else
                                ' ist nicht erlaubt
                                tmpValue = True
                            End If

                        End If

                    End If
                End If

            End If

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
                'Dim chckCell As Excel.Range = CType(appInstance.ActiveSheet.Cells(rng.Row, visboZustaende.meColpName), Excel.Range)

                'If CInt(chckCell.Interior.ColorIndex) <> XlColorIndex.xlColorIndexNone Then
                '    result = False
                'Else

                'End If
                ' tk, 16.9.18 das war vorher in dem Else-Zweig 
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
        'Dim zeilenWerte() As Double
        'Dim zeilensumme As Double
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
        'Dim roleCollection As New Collection
        'Dim roleUID As Integer
        Dim auslastungsArray(,) As Double = Nothing

        If isRole Then
            'roleUID = RoleDefinitions.getRoledef(rcName).UID
            'roleCollection.Add(rcName)
            'If awinSettings.meExtendedColumnsView Then
            '    auslastungsArray = visboZustaende.getUpDatedAuslastungsArray(roleCollection, von, bis, awinSettings.mePrzAuslastung)
            'End If

        End If


        Dim ixZeitraum As Integer
        Dim ix As Integer
        Dim breite As Integer
        Call awinIntersectZeitraum(phStart, phEnd, ixZeitraum, ix, breite)

        schnittmenge = calcArrayIntersection(von, bis, phStart, phEnd, xWerte)

        With CType(appInstance.ActiveSheet, Excel.Worksheet)
            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
        End With

        If schnittmenge.Sum > 0 Then
            For l As Integer = 0 To bis - von

                If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                    editRange.Cells(1, l + 1).value = schnittmenge(l)
                Else
                    editRange.Cells(1, l + 1).value = ""
                End If

            Next
        Else
            editRange.Value = ""
        End If


        appInstance.EnableEvents = formerEE

    End Sub

    ''' <summary>
    ''' markiert die angegebene Zeile, z.B zeichnet einen Rahmen drum herum 
    ''' </summary>
    ''' <param name="zeile"></param>
    Private Sub markZeile(ByVal zeile As Integer)

        Dim zRange As Excel.Range = Nothing

        With CType(appInstance.ActiveSheet, Excel.Worksheet)
            zRange = CType(.Range(.Cells(zeile, 1), .Cells(zeile, columnEndData)), Excel.Range)
        End With

        With zRange
            .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlDiagonalDown).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlDiagonalUp).LineStyle = XlLineStyle.xlLineStyleNone

            .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeLeft).Color = visboFarbeNone
            .Borders(XlBordersIndex.xlEdgeLeft).TintAndShade = 0
            .Borders(XlBordersIndex.xlEdgeLeft).Weight = XlBorderWeight.xlMedium

            .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeRight).Color = visboFarbeNone
            .Borders(XlBordersIndex.xlEdgeRight).TintAndShade = 0
            .Borders(XlBordersIndex.xlEdgeRight).Weight = XlBorderWeight.xlMedium

            .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeTop).Color = visboFarbeNone
            .Borders(XlBordersIndex.xlEdgeTop).TintAndShade = 0
            .Borders(XlBordersIndex.xlEdgeTop).Weight = XlBorderWeight.xlMedium

            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeBottom).Color = visboFarbeNone
            .Borders(XlBordersIndex.xlEdgeBottom).TintAndShade = 0
            .Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlMedium

        End With




    End Sub

    ''' <summary>
    ''' nimmt die Markeirung der Zeile wieder zurück 
    ''' </summary>
    ''' <param name="zeile"></param>
    Private Sub unMarkZeile(ByVal zeile As Integer)
        Dim zRange As Excel.Range = Nothing

        With CType(appInstance.ActiveSheet, Excel.Worksheet)
            zRange = CType(.Range(.Cells(zeile, 1), .Cells(zeile, columnEndData)), Excel.Range)
        End With

        With zRange
            .Borders(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlDiagonalDown).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlDiagonalUp).LineStyle = XlLineStyle.xlLineStyleNone

            .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlLineStyleNone
            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlLineStyleNone
        End With


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
