
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel

Public Class Tabelle2

    Private columnStartData As Integer = 8
    Private columnEndData As Integer = 30
    Private columnRC As Integer = 5
    Private oldValue As String = ""
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private meWS As Excel.Worksheet


    Private Sub Tabelle2_ActivateEvent() Handles Me.ActivateEvent

        
        Application.DisplayFormulaBar = False

        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        meWS = CType(appInstance.ActiveSheet, Excel.Worksheet)

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


        With Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .DisplayWorkbookTabs = False
            .GridlineColor = RGB(220, 220, 220)
            .FreezePanes = True
            '.DisplayHeadings = True
            .DisplayHeadings = False
        End With

        With meWS
            CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
        End With

        If Not IsNothing(appInstance.ActiveCell) Then
            oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If

        Application.EnableEvents = formerEE
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

        Try
            Dim auslastungChanged As Boolean = False
            Dim newValue As String = ""

            If Target.Column = columnRC Then
                ' es handelt sich um eine Änderung in der Rollen-/Kostenzuordnung
                If Target.Cells.Count = 1 Then
                    ' es handelt sich um eine Rollen- oder Kosten-Änderung ...
                    Dim zeile As Integer = Target.Row

                    newValue = CStr(Target.Cells(1, 1).value)
                    If isValidRCChange(newValue, oldValue) Then
                        ' es ist eine gültige Änderung, das heisst es wurde eine Rolle in eine andere gewechselt , oder 
                        ' eine Kostenart in eine andere; Kategorie-übergreifende Wechsel sind nicht erlaubt 

                        Dim pName As String = CStr(meWS.Cells(zeile, 2).value)
                        Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
                        Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
                        Dim phaseNameID As String = calcHryElemKey(phaseName, False)
                        Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
                        If Not IsNothing(curComment) Then
                            phaseNameID = curComment.Text
                        End If

                        ' jetzt muss noch geprüft werden, ob auch keine Duplikate vorkommen: zu einem Projekt dürfen z.Bsp keine 
                        ' 2 Zeilen existieren mit jeweils der gleichen Rolle oder Kostenart ...
                        If noDuplicatesInSheet(pName, phaseNameID, newValue, zeile) Then

                            Dim hproj As clsProjekt = ShowProjekte.getProject(pName)


                            If Not IsNothing(hproj) Then
                                Dim cPhase As clsPhase = hproj.getPhaseByID(phaseNameID)

                                If Not IsNothing(cPhase) Then
                                    If RoleDefinitions.containsName(newValue) Then
                                        ' es handelt sich um eine Rollen-Änderung
                                        If oldValue.Length > 0 And oldValue.Trim <> newValue.Trim Then
                                            ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                            Try
                                                Dim newRoleID As Integer = RoleDefinitions.getRoledef(newValue).UID
                                                cPhase.getRole(oldValue).RollenTyp = newRoleID
                                                auslastungChanged = True
                                            Catch ex As Exception
                                                Dim a As Integer = 0
                                            End Try

                                        Else
                                            ' es kam eine neue Rolle hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
                                            ' muss an dieser Stelle noch gar nichts gemacht werden ..
                                        End If
                                    Else
                                        ' es handelt sich um eine Kostenart Änderung 
                                        If oldValue.Length > 0 And oldValue.Trim <> newValue.Trim Then
                                            ' es handelt sich um einen Wechsel, von RoleID1 -> RoleID2
                                            Dim newCostID As Integer = CostDefinitions.getCostdef(newValue).UID
                                            cPhase.getCost(oldValue).KostenTyp = newCostID
                                        Else
                                            ' es kam eine neue Rolle hinzu, da es aber nicht möglich ist, im Datenbereich Eingaben zu machen, ohne dass eine Rolle / Kostenart ausgewählt wurde,
                                            ' muss an dieser Stelle noch gar nichts gemacht werden ..
                                        End If
                                    End If
                                End If

                            End If



                        Else
                            Call MsgBox("keine Doppelbelegung innerhalb einer Projektphase erlaubt ... ")
                            Target.Cells(1, 1).value = oldValue
                        End If
                        


                    Else
                        Call MsgBox("bitte nur innerhalb Rollen bzw. innerhalb Kostenarten wechseln !")
                        Target.Cells(1, 1).value = oldValue
                    End If
                Else
                    Call MsgBox("bitte nur eine Zelle selektieren ...")
                    Target.Cells(1, 1).value = oldValue
                End If

            Else

                ' es handelt sich um eine Datenänderung
                Call MsgBox("oldValue: " & oldValue)

            End If

            If auslastungChanged Then
                Dim roleNames As New Collection
                roleNames.Add(newValue, newValue)
                If oldValue.Length > 0 Then
                    If Not roleNames.Contains(oldValue) Then
                        roleNames.Add(oldValue, oldValue)
                    End If
                End If

                Call updateMassEditAuslastungsValues(showRangeLeft, showRangeRight, roleNames)
            End If

            oldValue = CStr(Target.Cells(1, 1).value)

        Catch ex As Exception
            Dim a As Integer = 0
        End Try
        
        appInstance.EnableEvents = True
    End Sub


    Private Sub Tabelle2_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange

        appInstance.EnableEvents = False

        If awinSettings.meEnableSorting Then
            ' es können auch nicht zugelassene Zellen selektiert worden sein 
            If Target.Cells.Count = 1 Then

                If isValidSelection(Target) Then
                    oldColumn = Target.Column
                    oldRow = Target.Row
                    If Not IsNothing(Target.Value) Then
                        oldValue = CStr(Target.Value)
                    Else
                        oldValue = ""
                    End If
                Else
                    CType(appInstance.ActiveSheet.Cells(oldRow, oldColumn), Excel.Range).Select()
                End If
                

            Else
                If isValidSelection(CType(Target.Cells(1, 1), Excel.Range)) Then
                    oldColumn = Target.Column
                    oldRow = Target.Row
                    If Not IsNothing(CType(Target.Cells(1, 1), Excel.Range).Value) Then
                        oldValue = CStr(CType(Target.Cells(1, 1), Excel.Range).Value)
                    Else
                        oldValue = ""
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
                oldValue = CStr(CType(Target.Cells(1, 1), Excel.Range).Value)
            Else
                oldValue = ""
            End If

            If Target.Column = columnRC Then
                'Call MsgBox("RoleCost")
            Else
                'Call MsgBox("Data")
            End If

        End If

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

        isValidSelection = result

    End Function

    ''' <summary>
    ''' prüft ob in dem aktiven Massen-Edit Sheet die übergebene Kombination nocheinmal vorkommt ... 
    ''' wenn nein: Rückgabe true
    ''' wenn ja: Rückgabe false
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcName"></param>
    ''' <param name="zeile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function noDuplicatesInSheet(ByVal pName As String, ByVal phaseNameID As String, ByVal rcName As String, _
                                             ByVal zeile As Integer) As Boolean
        Dim found As Boolean = False
        Dim curZeile As Integer = 2

        Dim chckName As String
        Dim chckPhNameID As String
        Dim chckRCName As String

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)

        With meWS
            chckName = CStr(meWS.Cells(curZeile, 2).value)

            Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
            chckPhNameID = calcHryElemKey(phaseName, False)
            Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
            If Not IsNothing(curComment) Then
                chckPhNameID = curComment.Text
            End If

            chckRCName = CStr(meWS.Cells(curZeile, 5).value)

        End With

        Do While Not found And curZeile <= visboZustaende.meMaxZeile


            If chckName = pName And _
                phaseNameID = chckPhNameID And _
                rcName = chckRCName And _
                zeile <> curZeile Then
                found = True
            Else
                curZeile = curZeile + 1

                With meWS
                    chckName = CStr(meWS.Cells(curZeile, 2).value)

                    Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                    chckPhNameID = calcHryElemKey(phaseName, False)
                    Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                    If Not IsNothing(curComment) Then
                        chckPhNameID = curComment.Text
                    End If

                    chckRCName = CStr(meWS.Cells(curZeile, 5).value)

                End With

            End If

        Loop

        noDuplicatesInSheet = Not found

    End Function
End Class
