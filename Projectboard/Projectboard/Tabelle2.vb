
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel

Public Class Tabelle2

    Private columnStartData As Integer = 8
    Private columnEndData As Integer = 30
    Private columnRC As Integer = 5
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private meWS As Excel.Worksheet


    Private Sub Tabelle2_ActivateEvent() Handles Me.ActivateEvent

        
        Application.DisplayFormulaBar = False

        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        'meWS = CType(appInstance.ActiveSheet, Excel.Worksheet)
        meWS = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(5)), Excel.Worksheet)

        Try
            ' die Anzahl maximaler Zeilen bestimmen 
            With visboZustaende
                .meMaxZeile = CType(appInstance.ActiveSheet, Excel.Worksheet).UsedRange.Rows.Count
                .meColRC = CType(appInstance.ActiveSheet.Range("RoleCost"), Excel.Range).Column
                .meColSD = CType(appInstance.ActiveSheet.Range("StartData"), Excel.Range).Column
                .meColED = CType(appInstance.ActiveSheet.Range("EndData"), Excel.Range).Column

                columnRC = .meColRC
                columnStartData = .meColSD
                columnEndData = .meColED
            End With
            
        Catch ex As Exception
            Call MsgBox("Fehler in Laden des Sheets ...")
        End Try
        
        Try
            If awinSettings.meEnableSorting Then
                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    .Unprotect("x")
                    .EnableSelection = XlEnableSelection.xlNoRestrictions
                End With
            Else
                With meWS
                    .Protect(Password:="x", UserInterfaceOnly:=True, _
                             AllowFormattingCells:=True, _
                             AllowInsertingColumns:=False,
                             AllowInsertingRows:=True, _
                             AllowDeletingColumns:=False, _
                             AllowDeletingRows:=True, _
                             AllowSorting:=True, _
                             AllowFiltering:=True)
                    .EnableSelection = XlEnableSelection.xlUnlockedCells
                    .EnableAutoFilter = True
                End With
            End If
            

        Catch ex As Exception

        End Try

        Try
            With Application.ActiveWindow
                .SplitColumn = 0
                .SplitRow = 1
                .DisplayWorkbookTabs = False
                .GridlineColor = RGB(220, 220, 220)
                .FreezePanes = True
                '.DisplayHeadings = True
                .DisplayHeadings = False
            End With

        Catch ex As Exception
            Call MsgBox("Fehler bei Activate Sheet Massen-Edit" & vbLf & ex.Message)
        End Try
        
        With meWS
            CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
        End With

        If Not IsNothing(appInstance.ActiveCell) Then
            visboZustaende.oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If

        Application.EnableEvents = True
        'Application.ScreenUpdating = True

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
            Dim newStrValue As String = ""

            Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
            Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(5)), Excel.Worksheet)

            If Target.Cells.Count = 1 Then

                Dim roleCostNames As New Collection

                Dim zeile As Integer = Target.Row
                Dim pName As String = CStr(meWS.Cells(zeile, 2).value)
                Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
                Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
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
                                    Dim parentName As String = RoleDefinitions.getParentRoleOf(tmprole.name)
                                    If parentName = "" Then
                                        rcValidation(i) = "alleRollen"
                                    Else
                                        rcValidation(i) = parentName
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
                                                cPhase.getRole(visboZustaende.oldValue).RollenTyp = newRoleID
                                                auslastungChanged = True
                                            Catch ex As Exception
                                                Dim a As Integer = 0
                                            End Try

                                        Else
                                            ' es kam eine neue Rolle hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
                                            ' muss an dieser Stelle nur die  gar nichts gemacht werden ..

                                        End If

                                        ' jetzt für die Zelle die Validation neu bestimmen, dazu muss aber der Blattschutz aufgehoben sein ...  

                                        If Not awinSettings.meEnableSorting Then
                                            ' es muss der Blattschutz aufgehoben werden, nachher wieder mit diesen Einstellungen aktiviert werden ...
                                            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                                .UnProtect(Password:="x")
                                            End With
                                        End If

                                        With currentCell

                                            Try
                                                If Not IsNothing(.Validation) Then
                                                    .Validation.Delete()
                                                End If
                                                ' jetzt wird die ValidationList aufgebaut 
                                                Dim tmpVal As String = validationStrings.Item(rcValidation(newRoleID))

                                                .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
                                                                               Formula1:=tmpVal)
                                            Catch ex As Exception

                                            End Try
                                        End With

                                        If Not awinSettings.meEnableSorting Then
                                            ' es muss der Blattschutz aufgehoben werden, nachher wieder mit diesen Einstellungen aktiviert werden ...
                                            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                                .Protect(Password:="x", UserInterfaceOnly:=True, _
                                                         AllowFormattingCells:=True, _
                                                         AllowInsertingColumns:=False,
                                                         AllowInsertingRows:=True, _
                                                         AllowDeletingColumns:=False, _
                                                         AllowDeletingRows:=True, _
                                                         AllowSorting:=True, _
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
                                            cPhase.getCost(visboZustaende.oldValue).KostenTyp = newCostID
                                        Else
                                            ' es kam eine neue Rolle hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
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
                        End If



                    Else
                        Call MsgBox("bitte nur innerhalb Rollen bzw. innerhalb Kostenarten wechseln !")
                        Target.Cells(1, 1).value = visboZustaende.oldValue
                    End If


                Else

                    ' es handelt sich um eine Datenänderung
                    Dim newDblValue As Double
                    Dim difference As Double

                    ' zu welcher / welchen Sammelrollen gehört die ausgewählte Rolle ? 
                    Dim sammelRollenName As String = ""
                    Dim zeileSammelRolle As Integer = 0
                    Dim isRole As Boolean

                    Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)
                    If RoleDefinitions.containsName(rcName) Then
                        isRole = True
                        ' hier muss jetzt bestimmt werden, wo die zugehörige Sammelrolle steht ... 
                    End If

                    If isRole Or CostDefinitions.containsName(rcName) Then
                        ' hier ist etwas gültiges vorhanden .. es kann also weitergemacht werden 

                        Try
                            newDblValue = CDbl(Target.Cells(1, 1).value)
                        Catch ex As Exception
                            newDblValue = 0.0
                        End Try

                        Try
                            Dim tmpWert As Double = CDbl(visboZustaende.oldValue)
                            difference = newDblValue - tmpWert
                        Catch ex As Exception
                            difference = newDblValue
                        End Try

                        Dim monthCol As Integer = showRangeLeft + CInt(((Target.Column - columnStartData) / 2))

                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)

                        If Not IsNothing(hproj) Then
                            Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)

                            If Not IsNothing(cphase) Then

                                Dim xWerteIndex As Integer = monthCol - getColumnOfDate(cphase.getStartDate)
                                Dim xWerteIndexChck As Integer = monthCol - (hproj.Start + cphase.relStart - 1)

                                Dim xWerte() As Double
                                Dim tmpSum As Double

                                If xWerteIndex <> xWerteIndexChck Then
                                    Call MsgBox("Kontrolle ... in Change Werte: " & xWerteIndex & ", " & _
                                                 xWerteIndexChck)
                                End If

                                If isRole Then
                                    ' es handelt sich um eine gültige Rolle

                                    If awinSettings.meAutoReduce Then
                                        'If awinSettings.meAutoReduce And difference > 0 Then
                                        ' nur dann muss die Sammelrolle entsprechend automatisch reduziert werden ... 

                                        Dim zeileOFSummaryRole As Integer = findeSammelRollenZeile(pName, phaseNameID, rcName)
                                        If zeileOFSummaryRole >= 2 And zeileOFSummaryRole <= visboZustaende.meMaxZeile Then
                                            Dim parentRoleName As String = CStr(meWS.Cells(zeileOFSummaryRole, columnRC).value)
                                            ' jetzt muss in der Sammel-Rolle aktualisiert werden 
                                            Dim parentRole As clsRolle = cphase.getRole(parentRoleName)

                                            If IsNothing(parentRole) Then
                                                ' nichts tun 
                                            Else
                                                ' der Monatswert muss geändert werden 
                                                xWerte = parentRole.Xwerte
                                                If xWerteIndex >= 0 And xWerteIndex <= xWerte.Length - 1 Then
                                                    Dim alterWert As Double = xWerte(xWerteIndex)
                                                    xWerte(xWerteIndex) = xWerte(xWerteIndex) - difference
                                                    If xWerte(xWerteIndex) < 0 Then
                                                        xWerte(xWerteIndex) = 0
                                                    End If
                                                    ' die Monatszahl und dann die Summe updaten ... 
                                                    CType(meWS.Cells(zeileOFSummaryRole, Target.Column), Excel.Range).Value = xWerte(xWerteIndex)

                                                    tmpSum = CDbl(CType(meWS.Cells(zeileOFSummaryRole, columnRC + 1), Excel.Range).Value)
                                                    tmpSum = tmpSum - System.Math.Min(alterWert, difference)
                                                    CType(meWS.Cells(zeileOFSummaryRole, columnRC + 1), Excel.Range).Value = tmpSum

                                                    summenChanged = True
                                                Else
                                                    Call MsgBox("Fehler in Übernahme Daten-Wert ...")
                                                End If
                                            End If
                                        End If

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
                                        xWerte(xWerteIndex) = newDblValue
                                        summenChanged = True
                                    Else
                                        Call MsgBox("Fehler in Übernahme Daten-Wert ...")
                                    End If

                                    tmpSum = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)
                                    tmpSum = tmpSum + difference
                                    CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = tmpSum

                                    ' bestimmt zu welchen Rollen die Auslastungs-Werte neu berechnet werden müssen ..
                                    roleCostNames = RoleDefinitions.getSummaryRoles(rcName)
                                    If Not roleCostNames.Contains(rcName) Then
                                        roleCostNames.Add(rcName, rcName)
                                    End If

                                    auslastungChanged = True


                                Else
                                    ' es handelt sich um eine gültige Kostenart - weiter oben wurde ja schon bestimmt, dass es entweder eine 
                                    ' gültige Rolle oder Kotenart ist 

                                    ' es muss einfach die Kostenart hinzugefügt bzw. die Werte abgeändert werden 
                                    Dim tmpCost As clsKostenart = cphase.getCost(rcName)

                                    If IsNothing(tmpCost) Then
                                        ' die Rolle muss neu angelegt und der Phase hinzugefügt werden  

                                        tmpCost = New clsKostenart(cphase.relEnde - cphase.relStart)
                                        tmpCost.KostenTyp = CostDefinitions.getCostdef(rcName).UID

                                        Call cphase.AddCost(tmpCost)

                                    End If

                                    ' der Monatswert muss geändert werden 
                                    xWerte = tmpCost.Xwerte
                                    If xWerteIndex >= 0 And xWerteIndex <= xWerte.Length - 1 Then
                                        xWerte(xWerteIndex) = newDblValue
                                        summenChanged = True
                                    Else
                                        Call MsgBox("Fehler in Übernahme Daten-Wert ...")
                                    End If

                                    ' jetzt die Summe neu ausgegeben ... 
                                    tmpSum = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)
                                    tmpSum = tmpSum + difference
                                    CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value = tmpSum

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


                If auslastungChanged Then
                    Call updateMassEditAuslastungsValues(showRangeLeft, showRangeRight, roleCostNames)
                End If

                ' das Folgende ist eigentlich eine Test Routine , die normalerweise gar nicht nötig ist 
                ' aber für Testzwecke gut geeignet ist ...

                Dim testValue1 As Double = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)
                If summenChanged Then
                    Call updateMassEditSummenValues(pName, phaseNameID, showRangeLeft, showRangeRight, roleCostNames)
                End If
                Dim testValue2 As Double = CDbl(CType(meWS.Cells(zeile, columnRC + 1), Excel.Range).Value)

                If testValue1 <> testValue2 Then
                    Call MsgBox("Unterschiede: " & testValue1 & ", " & testValue2)
                End If

                visboZustaende.oldValue = CStr(Target.Cells(1, 1).value)

            Else
                Call MsgBox("bitte nur eine Zelle selektieren ...")
                Target.Cells(1, 1).value = visboZustaende.oldValue
            End If


        Catch ex As Exception
            Call MsgBox("Fehler bei Massen-Edit, Ändern : " & vbLf & ex.Message)
        End Try
        
        appInstance.EnableEvents = True
    End Sub


    Private Sub Tabelle2_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange

        appInstance.EnableEvents = False

        Try
            ' wenn mehr wie eine Zelle selektiert wurde ...
            If Target.Cells.Count > 1 Then
                Target = CType(Target.Cells(1, 1), Excel.Range)
                Target.Select()
            End If

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
        End Try
        

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

        If RoleDefinitions.containsName(newValue) Then
            If RoleDefinitions.containsName(oldValue) Or oldValue = "" Then
                tmpValue = True
            End If
        ElseIf CostDefinitions.containsName(newValue) Then
            If CostDefinitions.containsName(oldValue) Or oldValue = "" Then
                tmpValue = True
            End If
        End If

        isValidRCChange = tmpValue

    End Function


    ''' <summary>
    ''' prüft, ob eine gültige Zelle selektiert wurde ... 
    ''' gültig ist eine Zelle, wenn sie entweder in der RoleCost Spalte ist oder in einer Datenspalte 
    ''' und ausserdem die Zeilennummer zwischen 2 und maxzeilen liegt 
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
                If rng.Row >= 2 And rng.Row <= visboZustaende.meMaxZeile Then
                    If rng.Column = columnRC Then
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

End Class
