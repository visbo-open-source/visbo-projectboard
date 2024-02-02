
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
    Private lastline As Integer = 2


    Private Sub Tabelle2_ActivateEvent() Handles Me.ActivateEvent


        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        Try

            Try
                Application.DisplayFormulaBar = False
                Application.ActiveWindow.DisplayWorkbookTabs = False
            Catch ex As Exception
                Call logger(ptErrLevel.logError, "DisplayFormularBar or DisplayWorkbookTybas failed ", ex.Message)
            End Try

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
                    'visboZustaende.meMaxZeile = CType(meWS, Excel.Worksheet).UsedRange.Rows.Count
                    visboZustaende.meColRC = CType(meWS.Range("RoleCost"), Excel.Range).Column
                    visboZustaende.meColSD = CType(meWS.Range("StartData"), Excel.Range).Column
                    visboZustaende.meColED = CType(meWS.Range("EndData"), Excel.Range).Column
                    visboZustaende.meColpName = 2
                    columnRC = .meColRC
                    columnStartData = .meColSD
                    columnEndData = .meColED
                    lastline = .meMaxZeile
                End With

            Catch ex As Exception
                Call MsgBox("Fehler in Laden des Sheets ...")
            End Try

            appInstance.EnableEvents = True
            Dim aa As Boolean = appInstance.EnableEvents

            ' jetzt die Spalte 6 einblenden bzw. ausblenden 
            Try
                If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                    CType(meWS.Columns("F"), Excel.Range).EntireColumn.Hidden = True
                    If editProjekteInSPE.Count = 1 Then
                        CType(meWS.Columns("A"), Excel.Range).Hidden = True
                        CType(meWS.Columns("B"), Excel.Range).Hidden = True
                        CType(meWS.Columns("C"), Excel.Range).Hidden = True
                    Else
                        CType(meWS.Columns("A"), Excel.Range).Hidden = False
                        CType(meWS.Columns("B"), Excel.Range).Hidden = False
                        CType(meWS.Columns("C"), Excel.Range).Hidden = False
                    End If
                ElseIf visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
                    If RoleDefinitions.getAllSkillIDs.Count > 0 Then
                        CType(meWS.Columns("F"), Excel.Range).Hidden = False
                    Else
                        CType(meWS.Columns("F"), Excel.Range).Hidden = True
                    End If
                    If editProjekteInSPE.Count = 1 Then
                        CType(meWS.Columns("A"), Excel.Range).Hidden = True
                        CType(meWS.Columns("B"), Excel.Range).Hidden = True
                        CType(meWS.Columns("C"), Excel.Range).Hidden = True
                    Else
                        CType(meWS.Columns("A"), Excel.Range).Hidden = False
                        CType(meWS.Columns("B"), Excel.Range).Hidden = False
                        CType(meWS.Columns("C"), Excel.Range).Hidden = False
                    End If

                End If
            Catch ex As Exception
                CType(meWS.Columns("F"), Excel.Range).Hidden = True
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
                        .EnableSelection = XlEnableSelection.xlNoRestrictions

                        meWS.EnableAutoFilter = True
                    End With
                End If


            Catch ex As Exception
                Call MsgBox("set autofilter in tabelle2 activateEvent")
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

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Error in ActivateEvent Tabelle meRC ", ex.Message)
        Finally

            If appInstance.EnableEvents = False Then
                appInstance.EnableEvents = True
            End If

        End Try

    End Sub

    Private Sub Tabelle2_BeforeDoubleClick(Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Me.BeforeDoubleClick

        Try
            If editProjekteInSPE.Count > 0 Then

                Dim former_EE As Boolean = appInstance.EnableEvents
                Dim newRangeLeft As Integer = 0
                Dim newRangeRight As Integer = 0

                appInstance.EnableEvents = True

                Dim currentCell As Excel.Range = Target

                ' die Rechtsklick-Behandlung soll auf alle Fälle abgeschaltet werden 
                Cancel = True

                Dim criteriaFulfilled As Boolean = False

                If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
                    criteriaFulfilled = (Target.Column = columnRC Or Target.Column = columnRC + 1) And (Target.Row > 1)
                End If

                If criteriaFulfilled Then

                    Try
                        Dim frmSelectCandidates As New frmAllocateRessources

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

                            If IsNothing(rcName) Then
                                rcName = ""
                            End If

                            If IsNothing(skillName) Then
                                skillName = ""
                            End If


                            If rcName <> "" Then

                                Dim rcNameID As String = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)


                                'Dim rcNameID As String = getRCNameIDfromExcelRange(CType(meWS.Range(Cells(zeile, columnRC), Cells(zeile, columnRC + 1)), Excel.Range))
                                Dim phaseNameID As String = getPhaseNameIDfromExcelCell(CType(meWS.Cells(zeile, columnRC - 1), Excel.Range))

                                Dim hproj As clsProjekt = Nothing
                                If Not IsNothing(pName) Then
                                    If pName <> "" Then
                                        hproj = ShowProjekte.getProject(pName)
                                    End If
                                End If

                                With frmSelectCandidates
                                    .hproj = hproj
                                    .rcNameID = rcNameID
                                    .phaseNameID = phaseNameID
                                End With

                                If Not IsNothing(hproj) And rcNameID <> "" And phaseNameID <> "" Then
                                    returnValue = frmSelectCandidates.ShowDialog()

                                    If returnValue = DialogResult.OK Then

                                        ' do the action 
                                        Dim myNewValue As Double = frmSelectCandidates.newValueForRCName
                                        ' now change the current amout for the summary role 
                                        Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)
                                        Dim myRole As clsRolle = cphase.getRoleByRoleNameID(rcNameID)


                                        Dim roleSkillValuesToAdd As SortedList(Of String, Double) = frmSelectCandidates.roleSkillValuesToAdd

                                        Dim startzeile As Integer = Target.Row
                                        Dim curZeile As Integer = startzeile
                                        Dim existingZeile As Integer = -1

                                        Dim phStart As Integer = hproj.Start + cphase.relStart - 1
                                        Dim phEnde As Integer = hproj.Start + cphase.relEnde - 1

                                        For Each kvp As KeyValuePair(Of String, Double) In roleSkillValuesToAdd

                                            Dim needNewZeile As Boolean = Not cphase.containsRoleSkillID(kvp.Key, inclChilds:=False, strictly:=True)
                                            If Not needNewZeile Then
                                                ' find the zeile containing project, phase, kvp.key: RoleID;SkillID
                                                existingZeile = findeZeileInMeRC(meWS, hproj.name, phaseNameID, kvp.Key)
                                            End If

                                            Dim ok As Boolean = cphase.substituteRole(rcNameID, kvp.Key, awinSettings.meAllowOverTime, kvp.Value)

                                            If ok Then
                                                Dim mytmpSkill As Integer = -1
                                                Dim myLoopRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(kvp.Key, mytmpSkill)
                                                Dim myLoopSkill As clsRollenDefinition = Nothing
                                                Dim myLoopSkillName As String = ""
                                                If mytmpSkill > 0 Then
                                                    myLoopSkill = RoleDefinitions.getRoleDefByID(mytmpSkill)
                                                    myLoopSkillName = myLoopSkill.name
                                                End If

                                                If needNewZeile Then
                                                    Call meRCZeileEinfuegen(zeile, myLoopRole.name, myLoopSkillName, True)
                                                    curZeile = curZeile + 1
                                                    zeile = visboZustaende.oldRow
                                                End If



                                                Dim mySubstituteRole As clsRolle = cphase.getRoleByRoleNameID(kvp.Key)

                                                If Not IsNothing(mySubstituteRole) AndAlso rcNameID <> kvp.Key Then

                                                    Dim tmpZeile As Integer = curZeile
                                                    If needNewZeile Then
                                                        tmpZeile = curZeile
                                                    Else
                                                        tmpZeile = existingZeile
                                                    End If

                                                    '' ur: 2022.03.29 
                                                    'Dim changed1 As Boolean = getTimeZoneRegardingTSO(newRangeLeft, newRangeRight, True)

                                                    '' aktualisiere die ergänzte Rolle 
                                                    'Call aktualisiereRoleCostInSheet(tmpZeile,
                                                    '                                     visboZustaende.meColSD, newRangeLeft, newRangeRight,
                                                    '                                     phStart, phEnde, mySubstituteRole.Xwerte)

                                                    'Call updateMassEditSummenValue(hproj, cphase, newRangeLeft, newRangeRight, kvp.Key, True, tmpZeile)

                                                    ' aktualisiere die ergänzte Rolle 
                                                    Call aktualisiereRoleCostInSheet(tmpZeile,
                                                                                         visboZustaende.meColSD, showRangeLeft, showRangeRight,
                                                                                         phStart, phEnde, mySubstituteRole.Xwerte)

                                                    Call updateMassEditSummenValue(hproj, cphase, showRangeLeft, showRangeRight, kvp.Key, True, tmpZeile)

                                                End If



                                            End If

                                        Next

                                        '' ur: 2022.03.29 
                                        'Dim changed2 As Boolean = getTimeZoneRegardingTSO(newRangeLeft, newRangeRight, True)

                                        '' aktualisiere die ursprüngliche Rolle 
                                        'Call aktualisiereRoleCostInSheet(Target.Row,
                                        '                                     visboZustaende.meColSD, newRangeLeft, newRangeLeft,
                                        '                                     phStart, phEnde, myRole.Xwerte)
                                        ' aktualisiere die ursprüngliche Rolle 
                                        Call aktualisiereRoleCostInSheet(Target.Row,
                                                                             visboZustaende.meColSD, showRangeLeft, showRangeRight,
                                                                             phStart, phEnde, myRole.Xwerte)

                                        ' den neuen Summenwert in die Summenspalte eintragen 
                                        Call updateMassEditSummenValue(hproj, cphase, showRangeLeft, showRangeRight, rcNameID, True, Target.Row)

                                        '

                                        '' aktualisieren der Charts 
                                        'Try

                                        '    If Not IsNothing(formProjectInfo1) Then
                                        '        Call updateProjectInfo1(visboZustaende.currentProject, visboZustaende.currentProjectinSession)
                                        '    End If
                                        '    ' tk 18.1.20
                                        '    Call aktualisiereCharts(visboZustaende.currentProject, True, calledFromMassEdit:=True, currentRCName:=rcName)

                                        '    Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcNameID)

                                        'Catch ex As Exception

                                        'End Try

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
                                            .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
                                            .EnableAutoFilter = True
                                        End With
                                    Else
                                        ' do nothing ...
                                    End If
                                End If

                            End If


                        End If

                    Catch ex As Exception
                        Call MsgBox("unexpected error 006A: " & ex.Message)

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
                            .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
                            .EnableAutoFilter = True
                        End With
                    End Try

                End If

                appInstance.EnableEvents = former_EE

            Else
                Dim msgTxt As String = "please load at least one project first"
                If Not awinSettings.englishLanguage Then
                    msgTxt = "bitte zunächst wenigstens ein Projekt laden"
                End If
                Call MsgBox(msgTxt)
            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Error in BeforeDoubleClick EventTabelle meRC ", ex.Message)
        End Try


    End Sub

    Private Sub Tabelle2_BeforeRightClick(Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Me.BeforeRightClick

        Try
            Dim msgTxt As String = ""
            If editProjekteInSPE.Count > 0 Then

                Dim former_EE As Boolean = appInstance.EnableEvents
                Dim addOrDeleteLine As New frmAddOrDeleteALine


                appInstance.EnableEvents = True

                Dim currentCell As Excel.Range = Target

                ' die Doubleklick-Behandlung soll auf alle Fälle abgeschaltet werden 
                Cancel = True

                Dim criteriaFulfilled As Boolean = False
                Dim criteriaFilterRequest As Boolean = ((Target.Row = 1) And (Target.Column = columnRC))
                Dim criteriaAddDeleteLine As Boolean = ((Target.Column > 0) And (Target.Column < columnRC) And (Target.Row > 1) And (Target.Row < visboZustaende.meMaxZeile))

                Dim isRole As Boolean = False

                If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
                    Try
                        isRole = True
                        criteriaFulfilled = (Target.Column = columnRC Or Target.Column = columnRC + 1) And (Target.Row > 1) And (CBool(Target.Locked) = False)
                    Catch ex As Exception

                    End Try


                ElseIf visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                    Try
                        criteriaFulfilled = (Target.Column = columnRC) And (Target.Row > 1) And (CBool(Target.Locked) = False)
                    Catch ex As Exception

                    End Try


                End If

                Dim zeile As Integer = Target.Row

                Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)


                Dim pname As String = ""
                Dim vName As String = ""
                Dim phaseName As String = ""
                Dim rcName As String = ""
                Dim skillName As String = ""
                Dim rcNameID As String = ""
                Dim phaseNameID As String = ""

                If zeile < visboZustaende.meMaxZeile Then

                    pname = CStr(meWS.Cells(zeile, visboZustaende.meColpName).value)
                    vName = CStr(meWS.Cells(zeile, 3).value)
                    phaseName = CStr(meWS.Cells(zeile, 4).value)
                    ' 
                    rcName = CStr(meWS.Cells(zeile, columnRC).value)
                    skillName = CStr(meWS.Cells(zeile, columnRC + 1).value)

                    If isRole Then
                        rcNameID = getRCNameIDfromExcelRange(CType(meWS.Range(Cells(zeile, columnRC), Cells(zeile, columnRC + 1)), Excel.Range))
                    Else
                        rcNameID = rcName
                    End If

                    phaseNameID = getPhaseNameIDfromExcelCell(CType(meWS.Cells(zeile, columnRC - 1), Excel.Range))

                End If



                ' prüfen, ob sich die selektierte Zelle in der Role-/Cost Spalte befindet 
                If criteriaFulfilled Then

                    Try

                        Dim frmMERoleCost As New frmMEhryRoleCost
                        Dim auslastungChanged As Boolean = False
                        Dim summenChanged As Boolean = False
                        ' muss extra überwacht werden, um das ProjectInfo1 Fenster auch immer zu aktualisieren
                        Dim kostenChanged As Boolean = False
                        Dim newStrValue As String = ""


                        Dim returnValue As DialogResult

                        If Target.Cells.Count = 1 Then

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

                            ' check, if the active ressource/role has some skills

                            Dim curRole As clsRollenDefinition = Nothing
                            Dim noSkills As Boolean = False

                            ' wenn rcname belegt ist, read roledefinition 
                            If rcName <> "" Then
                                curRole = RoleDefinitions.getRoledef(rcName)

                            End If

                            If Not IsNothing(curRole) Then

                                Dim topNodeSkills As List(Of Integer) = RoleDefinitions.getTopLevelTeamIDs
                                If Not curRole.isSummaryRole Then
                                    'check If the curRole does have skills
                                    noSkills = (curRole.getSkillCount <= 0)
                                Else
                                    ' es handelt sich um eine Summary Role 
                                    Dim ix As Integer = 1
                                    noSkills = True
                                    Do While ix <= topNodeSkills.Count And noSkills
                                        noSkills = (RoleDefinitions.getCommonChildsOfParents(topNodeSkills.ElementAt(ix - 1), curRole.UID).Count = 0)
                                        If noSkills Then
                                            ix = ix + 1
                                        End If
                                    Loop
                                End If

                            End If

                            If noSkills And frmMERoleCost.showSkillsOnly Then

                                If awinSettings.englishLanguage Then
                                    msgTxt = "Ressource " & rcName & " do not have any special skill"
                                Else
                                    msgTxt = "Ressource " & rcName & " hat keine speziellen skills"
                                End If
                                Call MsgBox(msgTxt)
                            Else

                                returnValue = frmMERoleCost.ShowDialog()

                                If returnValue = DialogResult.OK Then


                                    For Each roleSkillItem As String In frmMERoleCost.rolesToAdd
                                        Dim loopRcName As String = ""
                                        If frmMERoleCost.showSkillsOnly Then
                                            If rcName = "" Then

                                                Try
                                                    Dim tmpID As Integer = -1
                                                    loopRcName = RoleDefinitions.getContainingRoleOfSkillMembers(RoleDefinitions.getRoleDefByIDKennung(roleSkillItem, tmpID).UID).name

                                                    Dim chkRCNameID As String = RoleDefinitions.bestimmeRoleNameID(loopRcName, roleSkillItem)
                                                    If Not hproj.getPhaseByID(phaseNameID).containsRoleSkillID(chkRCNameID, inclChilds:=False, strictly:=True) Then
                                                        Call meRCZeileEinfuegen(zeile, loopRcName, roleSkillItem, True)
                                                        zeile = visboZustaende.oldRow
                                                    End If

                                                Catch ex As Exception

                                                End Try




                                            Else
                                                Try
                                                    Dim chkRCNameID As String = RoleDefinitions.bestimmeRoleNameID(loopRcName, roleSkillItem)
                                                    If Not hproj.getPhaseByID(phaseNameID).containsRoleSkillID(chkRCNameID, inclChilds:=False, strictly:=True) Then
                                                        Call meRCZeileEinfuegen(zeile, rcName, roleSkillItem, True)
                                                        zeile = visboZustaende.oldRow
                                                    End If
                                                Catch ex As Exception

                                                End Try

                                            End If

                                        Else
                                            Try
                                                Dim chkRCNameID As String = RoleDefinitions.bestimmeRoleNameID(roleSkillItem, skillName)
                                                If Not hproj.getPhaseByID(phaseNameID).containsRoleSkillID(chkRCNameID, inclChilds:=False, strictly:=True) Then
                                                    Call meRCZeileEinfuegen(zeile, roleSkillItem, skillName, True)
                                                    zeile = visboZustaende.oldRow
                                                End If
                                            Catch ex As Exception

                                            End Try


                                        End If



                                    Next

                                    For Each costNameIDitem As String In frmMERoleCost.costsToAdd
                                        Try
                                            Dim tmpCostID As Integer = CostDefinitions.getCostdef(costNameIDitem).UID
                                            If Not hproj.getPhaseByID(phaseNameID).containsCostID(tmpCostID) Then
                                                Call meRCZeileEinfuegen(zeile, costNameIDitem, "", False)
                                                zeile = visboZustaende.oldRow
                                            End If
                                        Catch ex As Exception

                                        End Try
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
                                        .EnableSelection = XlEnableSelection.xlNoRestrictions
                                        .EnableAutoFilter = True
                                    End With
                                    Cancel = True
                                End If


                            End If


                        Else
                            'Call MsgBox("bitte nur eine Zelle selektieren ...")
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
                            .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
                            .EnableAutoFilter = True
                        End With

                    End Try

                ElseIf criteriaFilterRequest = True Then

                    Try
                        Dim frmMERoleCost As New frmMEhryRoleCost
                        'Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)


                        If awinSettings.englishLanguage Then
                            frmMERoleCost.Text = "Select Name to filter Column"
                        Else
                            frmMERoleCost.Text = "Name selektieren um Spalte zu filtern"
                        End If

                        Dim returnValue As DialogResult = frmMERoleCost.ShowDialog()

                        If returnValue = DialogResult.OK Then

                            Dim ergebnisListe As New SortedList(Of String, String)

                            For Each roleSkillItem As String In frmMERoleCost.rolesToAdd
                                Try
                                    Dim curRoleSkill As clsRollenDefinition = RoleDefinitions.getRoledef(roleSkillItem)
                                    If Not IsNothing(curRoleSkill) Then

                                        If curRoleSkill.isSkill Then

                                            Dim childIds As SortedList(Of Integer, Double) = RoleDefinitions.getSubRoleIDsOf(curRoleSkill.name, type:=PTcbr.realRoles)
                                            For Each kvp As KeyValuePair(Of Integer, Double) In childIds
                                                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(kvp.Key)
                                                If Not ergebnisListe.ContainsKey(tmpRole.name) Then
                                                    ergebnisListe.Add(tmpRole.name, tmpRole.name)
                                                End If
                                            Next

                                        Else

                                            Dim childIds As SortedList(Of Integer, Double) = RoleDefinitions.getSubRoleIDsOf(curRoleSkill.name)
                                            For Each kvp As KeyValuePair(Of Integer, Double) In childIds
                                                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(kvp.Key)
                                                If Not ergebnisListe.ContainsKey(tmpRole.name) Then
                                                    ergebnisListe.Add(tmpRole.name, tmpRole.name)
                                                End If
                                            Next

                                        End If

                                    End If
                                Catch ex As Exception

                                End Try
                            Next


                            Dim ft As Array = ergebnisListe.Values.ToArray
                            Try

                                CType(meWS.Columns(columnRC), Excel.Range).AutoFilter(Field:=columnRC, Criteria1:=ft, [Operator]:=XlAutoFilterOperator.xlFilterValues)

                            Catch ex As Exception

                            End Try

                        End If



                    Catch ex As Exception
                        'Call MsgBox(ex.Message)
                    End Try

                Else
                    Try
                        ' Behandlung für Zeile hinzufügen/löschen
                        If criteriaAddDeleteLine = True Then

                            'Dim zeile As Integer = Target.Row
                            'Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                            'Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                            'Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)

                            addOrDeleteLine.position = Target
                            addOrDeleteLine.addLine = False
                            addOrDeleteLine.deleteLine = False
                            If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
                                addOrDeleteLine.isRoleSkill = True
                            ElseIf visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                                addOrDeleteLine.isCost = True
                            End If

                            addOrDeleteLine.isEmpty = rcName = ""

                            addOrDeleteLine.ShowDialog()
                            If addOrDeleteLine.addLine Then
                                ' Zeile hinzufügen
                                Call meRCZeileEinfuegen(zeile, "", "", isRole)

                            ElseIf addOrDeleteLine.deleteLine Then
                                ' Zeile löschen
                                Call meRCZeileLoeschen(zeile, pname, phaseNameID, rcNameID, isRole)
                            End If
                        End If
                    Catch ex As Exception

                    End Try

                End If

                appInstance.EnableEvents = former_EE

            Else
                msgTxt = "please load at least one project first"
                If Not awinSettings.englishLanguage Then
                    msgTxt = "bitte zunächst wenigstens ein Projekt laden"
                End If
                Call MsgBox(msgTxt)
            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Error in BeforeRightClick EventTabelle meRC ", ex.Message)
        Finally
            If appInstance.EnableEvents = False Then
                appInstance.EnableEvents = True
            End If
        End Try



    End Sub

    ''' <summary>
    ''' wird aufgerufen, sobald sich der Wert in einer Zelle verändert hat ...
    ''' entweder nachdem eine Dropbox Selection getroffen wurde oder eine Eingabe duch Pfeiltaste / Eingabe beendet wurde
    ''' 
    ''' </summary>
    ''' <param name="Target"></param>
    ''' <remarks></remarks>
    Private Sub Tabelle2_Change(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.Change

        Try
            If editProjekteInSPE.Count > 0 Then

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

                        Dim skillName As String = ""
                        If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
                            If Not IsNothing(meWS.Cells(zeile, columnRC + 1).value) Then
                                skillName = CStr(meWS.Cells(zeile, columnRC + 1).value).Trim
                            End If
                        End If


                        If Not IsNothing(meWS.Cells(zeile, columnRC).value) Then
                            rcName = CStr(meWS.Cells(zeile, columnRC).value).Trim
                            If rcName <> "" Then
                                isCost = CostDefinitions.containsName(rcName)
                                ' isRole wird erst später bestimmt, bleibt erst mal auf Falsch 
                            End If
                        End If

                        Dim phaseNameID As String = getPhaseNameIDfromExcelCell(CType(meWS.Cells(zeile, columnRC - 1), Excel.Range))

                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Dim cphase As clsPhase = Nothing


                        If Target.Columns.Count = 1 Then

                            If Not IsNothing(hproj) Then
                                cphase = hproj.getPhaseByID(phaseNameID)
                                If Not IsNothing(cphase) Then

                                    If Target.Column = columnRC Then
                                        ' es handelt sich um eine Rollen-Änderung ...

                                        Dim weitermachen As Boolean = True

                                        Dim skillID As Integer = -1
                                        Dim tryRcName As String = rcName

                                        If IsNothing(rcName) Then
                                            If Not IsNothing(visboZustaende.oldValue) Then
                                                If visboZustaende.oldValue <> "" And skillName = "" Then
                                                    Dim errMsg As String = "um Rolle /Kostenart zu löschen bitte entsprechenden Menupunkt nutzen ... "
                                                    If awinSettings.englishLanguage Then
                                                        errMsg = "to delete a role or cost, please use the according menu-item ..."
                                                    End If
                                                    Call MsgBox(errMsg)
                                                    Target.Cells(1, 1).value = visboZustaende.oldValue
                                                    weitermachen = False

                                                ElseIf skillName <> "" Then
                                                    Try
                                                        rcName = RoleDefinitions.getContainingRoleOfSkillMembers(RoleDefinitions.getRoledef(skillName).UID).name
                                                        Target.Cells(1, 1).value = rcName
                                                        tryRcName = rcName
                                                    Catch ex As Exception
                                                        Target.Cells(1, 1).value = visboZustaende.oldValue
                                                        weitermachen = False
                                                    End Try

                                                End If
                                            End If

                                        ElseIf rcName.Trim = "" Then

                                            If visboZustaende.oldValue <> "" And skillName = "" Then
                                                Dim errMsg As String = "um Rolle /Kostenart zu löschen bitte entsprechenden Menupunkt nutzen ... "
                                                If awinSettings.englishLanguage Then
                                                    errMsg = "to delete a role or cost, please use the according menu-item ..."
                                                End If
                                                Call MsgBox(errMsg)
                                                Target.Cells(1, 1).value = visboZustaende.oldValue
                                                weitermachen = False
                                            ElseIf skillName <> "" Then
                                                Try
                                                    rcName = RoleDefinitions.getContainingRoleOfSkillMembers(RoleDefinitions.getRoledef(skillName).UID).name
                                                    Target.Cells(1, 1).value = rcName
                                                    tryRcName = rcName
                                                Catch ex As Exception
                                                    Target.Cells(1, 1).value = visboZustaende.oldValue
                                                    weitermachen = False
                                                End Try
                                            End If

                                        End If

                                        If weitermachen Then

                                            If isValidRCChange(tryRcName, visboZustaende.oldValue, skillName, False) Then
                                                ' es ist eine gültige Änderung, das heisst es wurde eine Rolle in eine andere gewechselt , oder 
                                                ' eine Kostenart in eine andere; Kategorie-übergreifende Wechsel sind nicht erlaubt 

                                                ' jetzt muss noch geprüft werden, ob auch keine Duplikate vorkommen: zu einem Projekt dürfen z.Bsp keine 
                                                ' 2 Zeilen existieren mit jeweils der gleichen Rolle oder Kostenart ...

                                                ' jetzt ist 
                                                isCost = CostDefinitions.containsName(tryRcName)
                                                If Not isCost Then

                                                    If tryRcName <> rcName Then
                                                        ' tryRCName kann in isValidRCChange geändert worden sein 
                                                        rcName = tryRcName
                                                        meWS.Cells(zeile, columnRC).value = rcName
                                                    End If

                                                    ' isCOst ist falsch und isValidRCChange ... 
                                                    isRole = True

                                                    Dim autoDefineSkillName As Boolean = awinSettings.onePersonOneRole

                                                    If Not IsNothing(meWS.Cells(zeile, columnRC + 1).value) Then
                                                        skillName = CStr(meWS.Cells(zeile, columnRC + 1).value).Trim
                                                        If skillName.Length > 0 Then
                                                            If RoleDefinitions.containsName(skillName) Then
                                                                autoDefineSkillName = False
                                                                skillID = RoleDefinitions.getRoledef(skillName).UID
                                                            End If
                                                        End If

                                                    End If

                                                    If autoDefineSkillName And skillName = "" Then
                                                        Dim myRole As clsRollenDefinition = RoleDefinitions.getRoledef(rcName)
                                                        If Not IsNothing(myRole) Then
                                                            Dim mySkillIDS As SortedList(Of Integer, Double) = myRole.getSkillIDs()
                                                            If Not IsNothing(mySkillIDS) Then
                                                                If mySkillIDS.Count = 1 Then
                                                                    skillName = RoleDefinitions.getRoleDefByID(mySkillIDS.First.Key).name
                                                                    meWS.Cells(zeile, columnRC + 1).value = skillName
                                                                End If
                                                            End If
                                                        End If

                                                    End If

                                                    rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)

                                                ElseIf isCost Then

                                                    If tryRcName <> rcName Then
                                                        ' tryRCName kann in isValidRCChange geändert worden sein 
                                                        rcName = tryRcName
                                                        meWS.Cells(zeile, columnRC).value = rcName
                                                    End If

                                                    rcNameID = rcName


                                                End If


                                                If noDuplicatesInSheet(pName, phaseNameID, rcNameID, zeile) Then

                                                    Dim rcIndentLevel As Integer = 1
                                                    If isRole Then
                                                        ' es handelt sich um eine Rollen-Änderung


                                                        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(rcNameID, skillID)

                                                        ' jetzt den Indentlevel der Rolle vestimmen bestimmen 
                                                        'rcIndentLevel = RoleDefinitions.getRoleIndent(rcNameID)
                                                        currentCell.IndentLevel = rcIndentLevel

                                                        Dim newRoleID As Integer = tmpRole.UID
                                                        If visboZustaende.oldValue.Trim.Length > 0 And visboZustaende.oldValue.Trim <> rcName.Trim Then
                                                            ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                                            Try
                                                                auslastungChanged = True
                                                                Dim cRole As clsRolle = cphase.getRole(visboZustaende.oldValue, skillID)
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


                                                    ElseIf isCost Then

                                                        ' muss päter, wenn es Hierarchied er Kosten gibt ach angepasst werden. 
                                                        currentCell.IndentLevel = 1
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
                                                        ' tk 22.1.24 sollte dann nicht gelöscht werden 
                                                        'Call massEditZeileLoeschen("")

                                                    ElseIf RoleDefinitions.containsName(visboZustaende.oldValue) Then
                                                        Target.ClearComments()

                                                    End If


                                                End If

                                            Else
                                                Target.Cells(1, 1).value = visboZustaende.oldValue
                                            End If

                                        End If

                                    ElseIf Target.Column = columnRC + 1 Then
                                        ' es handelt sich um eine Skill Änderung

                                        Dim skillID As Integer = -1

                                        ' schon mal vorbelegen 
                                        If Not IsNothing(meWS.Cells(zeile, columnRC + 1).value) Then
                                            skillName = CStr(meWS.Cells(zeile, columnRC + 1).value).Trim
                                        End If

                                        Dim trySkillName As String = skillName
                                        If isValidRCChange(trySkillName, visboZustaende.oldValue, rcName, True) Or (skillName = "" And rcName <> "") Then

                                            If trySkillName <> skillName Then
                                                skillName = trySkillName
                                                meWS.Cells(zeile, columnRC + 1).value = skillName
                                            End If

                                            isRole = True

                                            Dim rcIndentLevel As Integer = 1
                                            If skillName <> "" Then
                                                'rcIndentLevel = RoleDefinitions.getRoleIndent(skillName)
                                                skillID = RoleDefinitions.getRoledef(skillName).UID
                                                Target.Cells(1, 1).IndentLevel = rcIndentLevel
                                            End If


                                            Dim rcNameGenerated As Boolean = False
                                            If rcName = "" Then
                                                rcName = RoleDefinitions.getContainingRoleOfSkillMembers(skillID).name
                                                rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                                                rcNameGenerated = True
                                            Else
                                                rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                                            End If

                                            If noDuplicatesInSheet(pName, phaseNameID, rcNameID, zeile) Then

                                                If rcNameGenerated Then
                                                    ' der automatisch generierte Name muss jetzt eingetragen werden 
                                                    Try
                                                        Target.Cells(1, 1).offset(0, -1).value = rcName
                                                        ' jetzt den richtigen Indent setzen ..
                                                        'rcIndentLevel = RoleDefinitions.getRoleIndent(rcName)
                                                        Target.Cells(1, 1).offset(0, -1).IndentLevel = rcIndentLevel
                                                    Catch ex As Exception

                                                    End Try

                                                End If


                                                Dim oldSkillID As Integer = -1
                                                If visboZustaende.oldValue <> "" Then
                                                    If RoleDefinitions.containsName(visboZustaende.oldValue) Then
                                                        oldSkillID = RoleDefinitions.getRoledef(visboZustaende.oldValue).UID
                                                    End If
                                                End If


                                                If oldSkillID <> skillID Then
                                                    ' es handelt sich um einen Wechsel 
                                                    Try
                                                        auslastungChanged = True
                                                        Dim cRole As clsRolle = cphase.getRole(rcName, oldSkillID)
                                                        If IsNothing(cRole) Then
                                                        Else
                                                            cRole.teamID = skillID
                                                        End If


                                                    Catch ex As Exception
                                                        visboZustaende.oldValue = ""

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
                                            Target.Cells(1, 1).value = visboZustaende.oldValue
                                        End If


                                    ElseIf Target.Column = columnRC + 2 Then
                                        ' es handelt sich um eine Summenänderung

                                        If Not IsNothing(meWS.Cells(zeile, columnRC).value) Then
                                            rcName = CStr(meWS.Cells(zeile, columnRC).value).Trim
                                            isRole = RoleDefinitions.containsName(rcName)
                                        End If

                                        If Not IsNothing(meWS.Cells(zeile, columnRC + 1).value) Then
                                            skillName = CStr(meWS.Cells(zeile, columnRC + 1).value).Trim
                                        End If

                                        If isRole Then
                                            rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                                        End If

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
                                            ok = False
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


                                                Dim von As Integer = showRangeLeft
                                                Dim bis As Integer = showRangeRight

                                                'Dim von As Integer = 0
                                                'Dim bis As Integer = 0
                                                '' ur: 2022.03.29 
                                                'Dim changed1 As Boolean = getTimeZoneRegardingTSO(von, bis, True)

                                                If isRole Then

                                                    ' jetzt muss die Rolle aktualisiert werden ...
                                                    Dim tmpRole As clsRolle = cphase.getRoleByRoleNameID(rcNameID)
                                                    Dim oldValues As Double()

                                                    ' calculate a distribution of values over months, dependent of months and number days / per Months
                                                    'considerValueOnly = True heisst, dass bei einem 1-dimensionaler
                                                    ' Xwerte Array die noNewCalculation, falls gesetzt, nicht berücksichtigt wird
                                                    Dim xValues() As Double = cphase.berechneBedarfeNew(xStartDate,
                                                                                                    xEndDate, vSum, 1, True)

                                                    If IsNothing(tmpRole) Then
                                                        ReDim oldValues(xValues.Length - 1)
                                                    Else
                                                        oldValues = tmpRole.Xwerte
                                                    End If

                                                    ' now check and verify whether this is feasible with given capacity 
                                                    ' if not, then do corrections in a way, that free capacity is taken and the rest of needs going over free capacity is distributed equally over the timeFrame
                                                    Dim allowOvertime As Boolean = awinSettings.meAllowOverTime
                                                    xValues = ShowProjekte.adjustToCapacity(uid, teamID, allowOvertime, xValues, xStartDate, oldValues)


                                                    If IsNothing(tmpRole) Then
                                                        tmpRole = New clsRolle(phEnde - phStart)

                                                        With tmpRole
                                                            .uid = uid
                                                            .teamID = teamID
                                                        End With
                                                        With cphase
                                                            .AddRole(tmpRole)
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

                                                    ' jetzt die tatsächliche Summe zeigen 
                                                    Target.Cells(1, 1).value = xValues.Sum

                                                    auslastungChanged = True

                                                    ' now, if requestedvalue vSum <> grantedValue xvalues.Sum 
                                                    ' provide feedback to user via comment
                                                    Try
                                                        meWS.Unprotect(Password:="x")

                                                        CType(Target.Cells(1, 1), Excel.Range).ClearComments()
                                                        Dim commentTxt As String = ""

                                                        If awinSettings.meAllowOverTime Then
                                                            commentTxt = "not assisted - this may cause overloads"
                                                        Else
                                                            commentTxt = getCommentTxt(vSum.Sum, xValues.Sum)
                                                        End If

                                                        CType(Target.Cells(1, 1), Excel.Range).AddComment(commentTxt)

                                                        With meWS
                                                            .Protect(Password:="x", UserInterfaceOnly:=True,
                                                                     AllowFormattingCells:=True,
                                                                     AllowFormattingColumns:=True,
                                                                     AllowInsertingColumns:=False,
                                                                     AllowInsertingRows:=False,
                                                                     AllowDeletingColumns:=False,
                                                                     AllowDeletingRows:=False,
                                                                     AllowSorting:=False,
                                                                     AllowFiltering:=True)
                                                        End With
                                                    Catch ex As Exception
                                                        With meWS
                                                            .Protect(Password:="x", UserInterfaceOnly:=True,
                                                                     AllowFormattingCells:=True,
                                                                     AllowFormattingColumns:=True,
                                                                     AllowInsertingColumns:=False,
                                                                     AllowInsertingRows:=False,
                                                                     AllowDeletingColumns:=False,
                                                                     AllowDeletingRows:=False,
                                                                     AllowSorting:=False,
                                                                     AllowFiltering:=True)
                                                        End With
                                                    End Try


                                                    Call aktualisiereRoleCostInSheet(Target.Row,
                                                                                 visboZustaende.meColSD, von, bis,
                                                                                 phStart, phEnde, tmpRole.Xwerte)


                                                Else

                                                    ' calculate a distribution of values over months, dependent of months and number days / per Months
                                                    'considerValueOnly = True heisst, dass bei einem 1-dimensionaler
                                                    ' Xwerte Array die noNewCalculation, falls gesetzt, nicht berücksichtigt wird
                                                    Dim xValues() As Double = cphase.berechneBedarfeNew(xStartDate,
                                                                                                    xEndDate, vSum, 1, True)

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
                                                    'Call aktualisiereRoleCostInSheet(Target.Row,
                                                    '                                     visboZustaende.meColSD, von, bis,
                                                    '                                     phStart, phEnde, xValues)
                                                    ' tk 4.1.24 last parameter has to have same dimension than whole phase
                                                    Call aktualisiereRoleCostInSheet(Target.Row,
                                                                                 visboZustaende.meColSD, von, bis,
                                                                                 phStart, phEnde, tmpCost.Xwerte)

                                                End If

                                            Else
                                                ' nichts tun 
                                            End If

                                        End If



                                    ElseIf Target.Column > columnRC + 2 Then
                                        ' es handelt sich um eine Datenänderung in den einzelnen Monaten 

                                        If Not IsNothing(meWS.Cells(zeile, columnRC).value) Then
                                            rcName = CStr(meWS.Cells(zeile, columnRC).value).Trim
                                            isRole = RoleDefinitions.containsName(rcName)
                                        End If

                                        If Not IsNothing(meWS.Cells(zeile, columnRC + 1).value) Then
                                            skillName = CStr(meWS.Cells(zeile, columnRC + 1).value).Trim
                                        End If

                                        If isRole Then
                                            rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                                        ElseIf isCost Then
                                            rcNameID = rcName
                                        End If



                                        ' zu welcher / welchen Sammelrollen gehört die ausgewählte Rolle ? 
                                        Dim sammelRollenName As String = ""
                                        Dim zeileSammelRolle As Integer = 0

                                        If isRole Or isCost Then
                                            ' hier ist etwas gültiges vorhanden .. es kann also weitergemacht werden 

                                            ' now check whether or not it is role and if it is valid input ..
                                            If isRole And Not awinSettings.meAllowOverTime Then

                                                Dim grantedValues As Double() = getGrantedValues(Target, pName, phaseNameID, rcNameID)

                                                Try
                                                    meWS.Unprotect(Password:="x")
                                                    Dim commentTxt As String = ""

                                                    For ix As Integer = 0 To grantedValues.Length - 1
                                                        ' so now correct input if necessary 
                                                        CType(Target.Cells(1, ix + 1), Excel.Range).ClearComments()

                                                        If IsNumeric(Target.Cells(1, ix + 1).value) Then

                                                            If awinSettings.meAllowOverTime Then
                                                                commentTxt = "not assisted - this may cause overloads"
                                                            Else
                                                                commentTxt = getCommentTxt(Target.Cells(1, ix + 1).value, grantedValues(ix))
                                                            End If

                                                            CType(Target.Cells(1, ix + 1), Excel.Range).AddComment(commentTxt)
                                                            Target.Cells(1, ix + 1).value = grantedValues(ix)

                                                        End If
                                                    Next

                                                    With meWS
                                                        .Protect(Password:="x", UserInterfaceOnly:=True,
                                                                 AllowFormattingCells:=True,
                                                                 AllowFormattingColumns:=True,
                                                                 AllowInsertingColumns:=False,
                                                                 AllowInsertingRows:=False,
                                                                 AllowDeletingColumns:=False,
                                                                 AllowDeletingRows:=False,
                                                                 AllowSorting:=False,
                                                                 AllowFiltering:=True)
                                                    End With

                                                Catch ex As Exception
                                                    With meWS
                                                        .Protect(Password:="x", UserInterfaceOnly:=True,
                                                                     AllowFormattingCells:=True,
                                                                     AllowFormattingColumns:=True,
                                                                     AllowInsertingColumns:=False,
                                                                     AllowInsertingRows:=False,
                                                                     AllowDeletingColumns:=False,
                                                                     AllowDeletingRows:=False,
                                                                     AllowSorting:=False,
                                                                     AllowFiltering:=True)
                                                    End With
                                                End Try


                                            End If

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

                            If Target.Column > columnRC + 2 Then
                                ' changes in months ... 

                                If Not IsNothing(meWS.Cells(zeile, columnRC).value) Then
                                    rcName = CStr(meWS.Cells(zeile, columnRC).value).Trim
                                    isRole = RoleDefinitions.containsName(rcName)
                                End If

                                If isRole Or isCost Then

                                    If isRole Then
                                        rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, skillName)
                                    ElseIf isCost Then
                                        rcNameID = rcName
                                    End If

                                    ' now check whether or not it is role and if it is valid input ..
                                    If isRole Then

                                        Dim grantedValues As Double() = getGrantedValues(Target, pName, phaseNameID, rcNameID)

                                        Try
                                            meWS.Unprotect(Password:="x")
                                            Dim commentTxt As String = ""
                                            For ix As Integer = 0 To grantedValues.Length - 1

                                                ' so now correct input if necessary 
                                                CType(Target.Cells(1, ix + 1), Excel.Range).ClearComments()

                                                If IsNumeric(Target.Cells(1, ix + 1).value) Then
                                                    If awinSettings.meAllowOverTime Then
                                                        commentTxt = "not assisted - this may cause overloads"
                                                    Else
                                                        commentTxt = getCommentTxt(Target.Cells(1, ix + 1).value, grantedValues(ix))
                                                    End If

                                                    CType(Target.Cells(1, 1), Excel.Range).AddComment(commentTxt)
                                                    Target.Cells(1, ix + 1).value = grantedValues(ix)

                                                End If

                                            Next
                                            With meWS
                                                .Protect(Password:="x", UserInterfaceOnly:=True,
                                                                     AllowFormattingCells:=True,
                                                                     AllowFormattingColumns:=True,
                                                                     AllowInsertingColumns:=False,
                                                                     AllowInsertingRows:=False,
                                                                     AllowDeletingColumns:=False,
                                                                     AllowDeletingRows:=False,
                                                                     AllowSorting:=False,
                                                                     AllowFiltering:=True)
                                            End With
                                        Catch ex As Exception
                                            With meWS
                                                .Protect(Password:="x", UserInterfaceOnly:=True,
                                                                     AllowFormattingCells:=True,
                                                                     AllowFormattingColumns:=True,
                                                                     AllowInsertingColumns:=False,
                                                                     AllowInsertingRows:=False,
                                                                     AllowDeletingColumns:=False,
                                                                     AllowDeletingRows:=False,
                                                                     AllowSorting:=False,
                                                                     AllowFiltering:=True)
                                            End With
                                        End Try

                                    End If

                                    Call updateDataValuesInProject(Target, isRole, rcName, rcNameID, pName, phaseNameID,
                                                                    auslastungChanged, summenChanged, kostenChanged, roleCostNames)
                                End If

                            End If


                        End If

                        If summenChanged Then

                            If IsNothing(cphase) Then
                                ' wenn in Zweig target.columns.count > 1 gewesen
                                cphase = hproj.getPhaseByID(phaseNameID)
                            End If

                            Call updateMassEditSummenValue(hproj, cphase, showRangeLeft, showRangeRight, rcNameID, isRole, zeile)

                        End If

                        If Not IsNothing(Target.Cells(1, 1).value) Then
                            visboZustaende.oldValue = CStr(Target.Cells(1, 1).value)
                        Else
                            visboZustaende.oldValue = ""
                        End If


                    ElseIf Target.Rows.Count > 1 Then

                        Dim errMsg As String = "Ändern der Ressourcen / Kosten nur innerhalb einer Zeile möglich !"
                        If awinSettings.englishLanguage Then
                            errMsg = "Editing resources / cost only is able in one line !"
                        End If
                        Call MsgBox(errMsg)

                        appInstance.ScreenUpdating = False
                        Call massEditRcTeAt(currentProjektTafelModus)

                        meWS.Activate()
                        appInstance.ScreenUpdating = True

                    End If


                Catch ex As Exception
                    Dim errMsg As String = "Fehler bei Massen-Edit, Ändern : " & vbLf & ex.Message
                    If awinSettings.englishLanguage Then
                        errMsg = "Error in editing resources / cost: " & vbLf & ex.Message
                    End If
                    Call MsgBox(errMsg)
                End Try



                appInstance.EnableEvents = True


                Else
                    If Not IsNothing(Target) Then
                    Try
                        If appInstance.EnableEvents = True Then
                            appInstance.EnableEvents = False
                        End If
                        CType(Target, Excel.Range).Clear()
                    Catch ex As Exception
                    Finally
                        If appInstance.EnableEvents = False Then
                            appInstance.EnableEvents = True
                        End If
                    End Try
                End If

                Dim msgTxt As String = "please load at least one project first"
                If Not awinSettings.englishLanguage Then
                    msgTxt = "bitte zunächst wenigstens ein Projekt laden"
                End If
                Call MsgBox(msgTxt)

            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Change Value Tabelle meRC ", ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' returns an array of values which are possible without causing bottlenecks
    ''' if there is no context provided: available capacity is considered to be the whole capacity of given rcNameID
    ''' if there is context provided:  available capacity is considered to be the remaining capacity under consideration of all other projects of the given context
    ''' </summary>
    ''' <param name="Target"></param>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcNameID"></param>
    ''' <returns></returns>
    Private Function getGrantedValues(ByVal Target As Excel.Range, ByVal pName As String,
                                      ByVal phaseNameID As String, ByVal rcNameID As String) As Double()

        Dim grantedValues As Double() ' holds finally the values which are granted
        Dim requestedValues As Double() ' holds the values as entered by user
        Dim projectValues As Double() ' holds the current values of roleNameID in phase/project

        If Not IsNothing(Target) Then

            Try
                Dim anzTargetColumns As Integer = Target.Columns.Count

                ReDim grantedValues(anzTargetColumns - 1)
                ReDim requestedValues(anzTargetColumns - 1)
                ReDim projectValues(anzTargetColumns - 1)


                ' now get the requestedValues from target 
                For ix As Integer = 0 To anzTargetColumns - 1
                    If IsNumeric(Target.Cells(1, ix + 1).value) Then
                        requestedValues(ix) = CDbl(Target.Cells(1, ix + 1).value)
                    Else
                        requestedValues(ix) = 0
                    End If
                Next

                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                If Not IsNothing(hproj) Then
                    Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)

                    If Not IsNothing(cphase) Then

                        Try

                        Catch ex As Exception

                        End Try

                        Dim roleUID As Integer
                        Dim teamID As Integer = -1

                        roleUID = RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID).UID

                        Dim von As Integer = showRangeLeft + Target.Column - columnStartData
                        Dim bis As Integer = von + anzTargetColumns - 1

                        ' freeCapacity now has the same dimension as requestedValues/target
                        Dim availableCapacity As Double() = ShowProjekte.getFreeCapacityOfRole(roleUID, teamID, von, bis)

                        Dim role As clsRolle = cphase.getRoleByRoleNameID(rcNameID)
                        If Not IsNothing(role) Then

                            ' you have to adjust the array availablecapacity by the projectvalues itself
                            Dim monthCol As Integer = showRangeLeft + Target.Column - columnStartData
                            Dim xWerteIndex As Integer = monthCol - getColumnOfDate(cphase.getStartDate)
                            Dim xWerte() As Double = role.Xwerte

                            For ix As Integer = 0 To anzTargetColumns - 1

                                Try
                                    projectValues(ix) = xWerte(xWerteIndex + ix)
                                    ' available capacity is now including projectvalues, because the project values are being substututed
                                    availableCapacity(ix) = availableCapacity(ix) + projectValues(ix)
                                Catch ex As Exception
                                    Call logger(ptErrLevel.logError, "getGrantedValues in Edit Cell " & pName & " " & phaseNameID & " " & rcNameID, ex.Message)
                                End Try

                            Next
                        End If


                        ' now do the Job ... 
                        For ix As Integer = 0 To anzTargetColumns - 1
                            grantedValues(ix) = System.Math.Min(requestedValues(ix), availableCapacity(ix))
                        Next


                    Else
                        Call logger(ptErrLevel.logError, "getGrantedValues in Edit Cell " & pName & " " & phaseNameID, " : Phase was Nothing ..")
                    End If
                Else
                    Call logger(ptErrLevel.logError, "getGrantedValues in Edit Cell " & pName, " : hproj was Nothing ..")
                End If
            Catch ex As Exception

            End Try

        Else
            ReDim grantedValues(0)
            Call logger(ptErrLevel.logError, "getGrantedValues in Edit Cell " & pName, " Cell was Nothing ..")
        End If


        getGrantedValues = grantedValues
    End Function
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

                ' tk: target.column - returns the nr of the first column of the range  
                Dim monthCol As Integer = showRangeLeft + target.Column - columnStartData

                Dim xWerteIndex As Integer = monthCol - getColumnOfDate(cphase.getStartDate)
                Dim xWerte() As Double

                If isRole Then
                    ' es handelt sich um eine gültige Rolle

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


                    auslastungChanged = True


                Else
                    ' es handelt sich um eine gültige Kostenart - weiter oben wurde ja schon bestimmt, dass es entweder eine 
                    ' gültige Rolle oder Kotenart ist 

                    ' es muss einfach die Kostenart hinzugefügt bzw. die Werte abgeändert werden 
                    Dim tmpCost As clsKostenart = cphase.getCost(rcName)

                    If IsNothing(tmpCost) Then
                        ' die Kostenart muss neu angelegt und der Phase hinzugefügt werden  

                        tmpCost = New clsKostenart(cphase.relEnde - cphase.relStart) With {
                            .KostenTyp = CostDefinitions.getCostdef(rcName).UID
                        }

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


                End If
            End If

        End If





    End Sub


    Private Sub Tabelle2_Deactivate() Handles Me.Deactivate

        Try
            appInstance.ActiveWindow.SplitColumn = 0
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWindow.SplitRow = 0
        Catch ex As Exception

        End Try
        Try
            Me.Columns.Hidden = False

        Catch ex As Exception

        End Try

        Try
            appInstance.DisplayFormulaBar = False
        Catch ex As Exception

        End Try

        ' tk 23.11.22 do not update screen ...
        ' is again activated in Activate event of each table 
        Try
            appInstance.ScreenUpdating = False
        Catch ex As Exception

        End Try


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

        If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
            CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart & pdEinheit

        ElseIf visboZustaende.projectBoardMode = ptModus.massEditCosts Then
            CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart & "[T€]"

        Else
            CType(meWS.Cells(1, visboZustaende.meColRC + 2), Excel.Range).Value = headerPart

        End If




    End Sub

    ''' <summary>
    ''' wird aufgerufen, wenn sich die Zeile ändert ...
    ''' </summary>
    ''' <param name="Target"></param>
    Private Sub Tabelle2_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange

        Try
            If editProjekteInSPE.Count > 0 Then
                appInstance.EnableEvents = False
                Dim former_showRangeLeft As Integer = showRangeLeft
                Dim former_showRangeRight As Integer = showRangeRight

                ' tk 23.1.24 now check whether you are at the lastLine+1 , then add new line 
                If Target.Row = visboZustaende.meMaxZeile Then
                    Dim newMaxZeile As Integer = visboZustaende.meMaxZeile
                    Dim isRole As Boolean = True
                    If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                        isRole = False
                    End If
                    Call meRCZeileEinfuegen(newMaxZeile - 1, "", "", isRole)
                    appInstance.EnableEvents = False
                End If

                Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)

                Try

                    If Target.Row <> oldRow Then
                        'Call highlightRow(Target.Row, oldRow)

                        ' jetzt muss in der Spaltenüberschrift noch angegeben werden, ob es sich um T€, PT oder nichts handelt 
                        Call defineHeaderTitleOfRoleCost(Target.Row)

                    End If



                Catch ex As Exception

                End Try



                Dim pname As String = ""
                Dim rcName As String = ""
                Dim rcNameID As String = ""
                Dim oldRCName As String = ""
                Dim oldRCNameID As String = ""

                Dim oldElemID As String = visboZustaende.currentElemID

                Dim changeBecauseRCNameChanged As Boolean = False
                Dim changeBecausePhaseNameIDChanged As Boolean = False
                Try
                    ' wenn mehr wie eine Zelle selektiert wurde ...
                    If Target.Cells.Count > 1 Then
                        Target = CType(Target.Cells(1, 1), Excel.Range)
                        Target.Select()
                    End If

                    rcName = CStr(meWS.Cells(Target.Row, columnRC).value)
                    rcNameID = getRCNameIDfromExcelRange(CType(meWS.Range(meWS.Cells(Target.Row, columnRC), meWS.Cells(Target.Row, columnRC + 1)), Excel.Range))


                    ' um welche Phase handelt es sich ? 
                    Try
                        visboZustaende.currentElemID = getPhaseNameIDfromExcelCell(CType(meWS.Cells(Target.Row, columnRC - 1), Excel.Range))
                        changeBecausePhaseNameIDChanged = (oldElemID <> visboZustaende.currentElemID)
                    Catch ex As Exception

                    End Try

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

                    If pNameChanged Or changeBecauseRCNameChanged Or (changeBecausePhaseNameIDChanged And Not awinSettings.considerProjectTotals) Then

                        ' umgesetzte timeZone
                        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff(True)

                        ' tk 16.11 
                        'Call aktualisiereCharts(.currentProject, True, calledFromMassEdit:=True, currentRCName:=rcName)

                        If pNameChanged Then
                            selectedProjekte.Clear(False)
                            selectedProjekte.Add(.currentProject, False)
                        End If

                        If Not IsNothing(rcNameID) Then

                            If rcNameID <> "" Then
                                Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcNameID)
                            End If
                        End If


                        'If Not IsNothing(formProjectInfo1) Then
                        '    Call updateProjectInfo1(.currentProject, .currentProjectinSession)
                        '    ' hier wird dann ggf noch das Projekt-/RCNAme/aktuelle Version vs DB-Version Chart aktualisiert  
                        'End If


                    End If


                End With
                appInstance.EnableEvents = True

                ' zurücksetzen timezone
                showRangeRight = former_showRangeRight
                showRangeLeft = former_showRangeLeft

            Else
                Dim msgTxt As String = "please load at least one project first"
                If Not awinSettings.englishLanguage Then
                    msgTxt = "bitte zunächst wenigstens ein Projekt laden"
                End If
                Call MsgBox(msgTxt)
            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Error in SelectionChange in EventTabelle meRC ", ex.Message)
        End Try

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
    Private Function isValidRCChange(ByRef newValue As String, ByVal oldValue As String, ByVal otherValue As String, ByVal isSkillCheck As Boolean) As Boolean

        Dim rcName As String = ""
        Dim skillName As String = ""
        Dim tmpValue As Boolean = False
        Dim msgTxt As String = ""

        If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then

            If isSkillCheck Then
                ' es handelt sich um den Skill Check 
                rcName = otherValue
                skillName = newValue

                ' hier muss die Prüfung rein , ob es eine bekannte Skill ist ...
                If Not RoleDefinitions.containsName(skillName) Then
                    Dim skillNamesList As List(Of String) = RoleDefinitions.getSkillNamesContainingSubStr(skillName, otherValue)
                    If skillNamesList.Count = 1 Then
                        skillName = CStr(skillNamesList.Item(0))
                        newValue = skillName
                    ElseIf skillNamesList.Count > 1 Then
                        ' Formular mit Liste zeigen 
                        Dim selectionFrm As New frmSelectOneItem
                        If awinSettings.englishLanguage Then
                            selectionFrm.Text = "Select one skill"
                        Else
                            selectionFrm.Text = "Wählen Sie eine Skill"
                        End If
                        selectionFrm.itemsCollection = skillNamesList
                        If selectionFrm.ShowDialog = DialogResult.OK Then
                            skillName = CStr(selectionFrm.itemList.SelectedItem)
                            newValue = skillName
                        End If
                    End If
                End If

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
                ' hier muss die Prüfung rein , ob es eine bekannte Kostenart ist ...
                ' hier muss die Prüfung rein , ob es eine bekannte Rolle ist ...
                If Not RoleDefinitions.containsName(rcName) Then
                    Dim roleNamesList As List(Of String) = RoleDefinitions.getRoleNamesContainingSubStr(rcName, otherValue)
                    If roleNamesList.Count = 1 Then
                        rcName = CStr(roleNamesList.Item(0))
                        newValue = rcName
                    ElseIf roleNamesList.Count > 1 Then
                        ' Formular mit Liste zeigen 
                        Dim selectionFrm As New frmSelectOneItem
                        If awinSettings.englishLanguage Then
                            selectionFrm.Text = "Select one Resource"
                        Else
                            selectionFrm.Text = "Wählen Sie eine Resource"
                        End If
                        selectionFrm.itemsCollection = roleNamesList
                        If selectionFrm.ShowDialog = DialogResult.OK Then
                            rcName = CStr(selectionFrm.itemList.SelectedItem)
                            newValue = rcName
                        End If
                    End If

                End If

                Dim stillOk As Boolean = RoleDefinitions.containsName(rcName)

                If Not stillOk Then
                    msgTxt = "unbekannt: " & rcName
                    If awinSettings.englishLanguage Then
                        msgTxt = "unknown: " & rcName
                    End If
                End If

                If skillName <> "" And stillOk Then
                    ' prüfen, ob es diese Skill denn überhaupt für die Rolle gibt 
                    ' wenn es sich um eine Kostenart handelt : sillok = false
                    stillOk = RoleDefinitions.roleHasSkill(rcName, skillName)
                    If Not stillOk Then
                        msgTxt = "passt nicht zu Skill " & skillName & ": " & rcName
                        If awinSettings.englishLanguage Then
                            msgTxt = "does not have appropriate skill: " & skillName & ": " & rcName
                        End If
                    End If
                End If

                If stillOk Then

                    Dim weiterMachen As Boolean = False
                    Dim skillID As Integer = -1

                    ' erstmal prüfen, ob es sich um einen Ressourcen-Manager oder Portfolio Manager handelt; denn dann können nicht alle Werte eingegeben werden 
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                        Dim parentCollection As New Collection From {
                            RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, skillID).name
                        }

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
                                Else
                                    ' ist nicht erlaubt
                                    tmpValue = False
                                End If
                            Else
                                tmpValue = False
                            End If

                        End If
                    End If
                Else
                    Call MsgBox(msgTxt)
                    tmpValue = False
                End If


            End If

        ElseIf visboZustaende.projectBoardMode = ptModus.massEditCosts Then
            rcName = newValue
            skillName = otherValue

            If Not CostDefinitions.containsName(rcName) Then
                Dim costNamesList As List(Of String) = CostDefinitions.getCostNamesContainingSubStr(rcName)
                If costNamesList.Count = 1 Then
                    rcName = CStr(costNamesList.Item(0))
                    newValue = rcName
                ElseIf costNamesList.Count > 1 Then
                    ' Formular mit Liste zeigen 
                    Dim selectionFrm As New frmSelectOneItem
                    If awinSettings.englishLanguage Then
                        selectionFrm.Text = "Select one Cost"
                    Else
                        selectionFrm.Text = "Wählen Sie eine Kostenart"
                    End If
                    selectionFrm.itemsCollection = costNamesList
                    If selectionFrm.ShowDialog = DialogResult.OK Then
                        rcName = CStr(selectionFrm.itemList.SelectedItem)
                        newValue = rcName
                    Else
                        rcName = ""
                        newValue = ""
                    End If
                End If
            End If

            If rcName <> "" Then
                If CostDefinitions.containsName(rcName) Then
                    tmpValue = True
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
    ''' aktualisiert die Werte in der angegebenen Zeile mit den Daten aus XWerte 
    ''' funktioniert sowohl für Rollen als auch Kosten 
    ''' </summary>
    ''' <param name="zeile"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="phStart">ist pStart+relstart-1</param>
    ''' <param name="phEnd">ist pStart+relende -1</param>
    ''' <param name="xWerte"></param>
    ''' <remarks></remarks>
    Private Sub aktualisiereRoleCostInSheet(ByVal zeile As Integer,
                                      ByVal startSpalteDaten As Integer,
                                      ByVal von As Integer, ByVal bis As Integer,
                                      ByVal phStart As Integer, ByVal phEnd As Integer,
                                      ByVal xWerte() As Double)
        Dim schnittmenge() As Double

        Dim editRange As Excel.Range

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' tk this leads to error if von bis does not correspond to Xwerte
        ' sicherstellen, dass die Länge von xWerte = phStart-phEnd +1 ist
        ' sonst funktioniert die Zuweisung weiter unten nicht 
        'If phStart < von Then
        '    phStart = von
        'End If
        'If phEnd > bis Then
        '    phEnd = bis
        'End If


        Dim ixZeitraum As Integer
        Dim ix As Integer
        Dim breite As Integer
        ' define breite, iXZeitraum and IX
        ' tk 8.1.24
        'Call awinIntersectZeitraum(von, bis, ixZeitraum, ix, breite)
        ' von ist gleich showrangeLeft, bis ist gleich showrangeRight
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
    ''' returns a commentTxt to be used in massEdit Resources, asking for more resources than are available
    ''' </summary>
    ''' <param name="requested"></param>
    ''' <param name="granted"></param>
    ''' <returns></returns>
    Private Function getCommentTxt(ByVal requested As Double, ByVal granted As Double) As String

        Dim contextTxt As String
        If projectConstellations.Count > 0 Then
            contextTxt = "Consider all " & projectConstellations.Liste.First.Key
            If contextTxt.Last = "#" Then
                contextTxt = contextTxt.Substring(0, contextTxt.Length - 1)
            End If
        Else
            contextTxt = "Consider no other projects"
        End If

        Dim commentTxt As String = contextTxt & vbLf & "requested: " & requested.ToString("0.#") & vbLf & "granted: " & granted.ToString("0.#")

        getCommentTxt = commentTxt
    End Function
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

    Private Sub Tabelle2_Startup(sender As Object, e As EventArgs) Handles Me.Startup
        If visboClient = divClients(client.VisboSPE) Then
            'Call MsgBox("bin im meRC")
        End If
    End Sub


    '''' <summary>
    '''' blendet die Spalte aus
    '''' </summary>
    '''' <param name="spalte"></param>
    'Private Sub ausblendenSpalte(ByVal spalte As Integer)
    '    Dim zRange As Excel.Range = Nothing

    '    With CType(appInstance.ActiveSheet, Excel.Worksheet)
    '        zRange = CType(.Range(.Cells(1, spalte), .Cells(lastline, spalte)), Excel.Range)
    '    End With

    '    Dim colSpalte As Range = zRange.EntireColumn.Hidden()

    'End Sub

End Class
