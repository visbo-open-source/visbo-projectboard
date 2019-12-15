
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel

Public Class Tabelle3

    Private columnStartDate As Integer = 5
    Private columnEndDate As Integer = 6

    Private oldColumn As Integer = 6
    Private oldRow As Integer = 2

    Private columnName As Integer = 2

    ' eine Enum, über die die Spalten adressiert werden können
    Private Enum PTmeTe
        projectNr = 0
        pName = 1
        vName = 2
        elemName = 3
        startdate = 4
        endDate = 5
        trafficLight = 6
        explanation = 7
        deliverables = 8
        responsible = 9
        percentDone = 10
        documentLink = 11
    End Enum

    ' enthält die Spalten, wo die einzelnen Felder stehen , korreliert mit der Enum allianzSpalten
    Private col() As Integer

    Private Sub Tabelle3_ActivateEvent() Handles Me.ActivateEvent

        ' in der Mass-Edit Termine sollen Header und Formular-Bar immer erhalten bleiben ...
        Application.DisplayFormulaBar = True

        Dim enumTermineColumnsCount As Integer = [Enum].GetNames(GetType(PTmeTe)).Length
        ReDim col(enumTermineColumnsCount)


        col(PTmeTe.projectNr) = 1
        col(PTmeTe.pName) = 2
        col(PTmeTe.vName) = 3
        col(PTmeTe.elemName) = 4
        col(PTmeTe.startdate) = 5
        col(PTmeTe.endDate) = 6
        col(PTmeTe.trafficLight) = 7
        col(PTmeTe.explanation) = 8
        col(PTmeTe.deliverables) = 9
        col(PTmeTe.responsible) = 10
        col(PTmeTe.percentDone) = 11
        col(PTmeTe.documentLink) = 12

        ' initial setzen der Spalten ... 

        'Dim filterRange As Excel.Range
        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meTE)), Excel.Worksheet)


        ' jetzt den Schutz aufheben , falls einer definiert ist 
        If meWS.ProtectContents Then
            meWS.Unprotect(Password:="x")
        End If

        Try
            ' die Anzahl maximaler Zeilen bestimmen 
            With visboZustaende
                .meMaxZeile = CType(meWS, Excel.Worksheet).UsedRange.Rows.Count
                ' ist die Spalte für MSTask-Name 
                .meColRC = 4
                ' ist die Spalte für Startdate
                .meColSD = 5
                ' ist die Spalte für Ende-Date
                .meColED = 6
                ' ist die Spalte für den Projekt-Namen 
                .meColpName = 2

                columnStartDate = .meColSD
                columnEndDate = .meColED
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
            ' es dürfen keine Zeilen ergänzt werden, noch Spalten 
            ' die dürfen auch nicht gelöscht werden 
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
                .EnableSelection = XlEnableSelection.xlUnlockedCells
                .EnableAutoFilter = True
            End With


        Catch ex As Exception

        End Try


        Application.EnableEvents = formerEE

        ' einen Select machen - nachdem Event Behandlung wieder true ist, dann werden project und lastprojectDB gesetzt ...
        Try
            Dim cz As Integer = 2
            Dim eof As Boolean = (cz > visboZustaende.meMaxZeile)

            Dim bedingung As Boolean = CBool(CType(meWS.Cells(cz, col(PTmeTe.trafficLight)), Excel.Range).Locked = True) And Not eof

            Do While bedingung
                cz = cz + 1
                eof = (cz > visboZustaende.meMaxZeile)
                bedingung = CBool(CType(meWS.Cells(cz, col(PTmeTe.trafficLight)), Excel.Range).Locked = True) And Not eof
            Loop

            If Not eof Then
                CType(CType(meWS, Excel.Worksheet).Cells(cz, col(PTmeTe.trafficLight)), Excel.Range).Select()

                Dim pName As String = ""

                With visboZustaende

                    pName = CStr(CType(meWS.Cells(cz, visboZustaende.meColpName), Excel.Range).Value)
                    If ShowProjekte.contains(pName) Then
                        .currentProject = ShowProjekte.getProject(pName)
                        .currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pName, .currentProject.variantName))
                    End If

                End With
            Else
                CType(CType(meWS, Excel.Worksheet).Cells(cz, col(PTmeTe.trafficLight)), Excel.Range).Locked = False
                CType(CType(meWS, Excel.Worksheet).Cells(cz, col(PTmeTe.trafficLight)), Excel.Range).Select()
            End If



        Catch ex As Exception

        End Try

        ' jetzt die Gridline zeigen
        With appInstance.ActiveWindow
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


    End Sub

    Private Sub Tabelle3_Change(Target As Range) Handles Me.Change

        ' damit nicht eine immerwährende Event Orgie durch Änderung in den Zellen abgeht ...
        appInstance.EnableEvents = False

        Dim currentCell As Excel.Range = Target
        Dim cphase As clsPhase = Nothing
        Dim cMilestone As clsMeilenstein = Nothing


        Dim hproj As clsProjekt = visboZustaende.currentProject

        If IsNothing(hproj) Then
            Call MsgBox("Projekt konnte nicht bestimmt werden ...")
            appInstance.EnableEvents = True
            Exit Sub
        Else

            Dim allowedLeftDate As Date = hproj.startDate
            Dim allowedRightDate As Date = hproj.endeDate

            Try
                Dim datesWereChanged As Boolean = False

                Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meTE)), Excel.Worksheet)

                If Target.Cells.Count = 1 Then

                    Dim currentZeile As Integer = Target.Row
                    Dim currentColumn As Integer = Target.Column

                    Dim elemID As String = visboZustaende.currentElemID

                    ' jetzt bestimmen, ob es sich bei dem Eintrag in der Zeile um eine Phase oder einen Meilenstein handelt
                    Dim elemIsMilestone As Boolean = elemIDIstMeilenstein(elemID)
                    If elemIsMilestone Then
                        cMilestone = hproj.getMilestoneByID(elemID)
                        cphase = Nothing
                    Else
                        cMilestone = Nothing
                        cphase = hproj.getPhaseByID(elemID)
                    End If

                    ' dann die allowdLeft und RightDate berechnen
                    ' jedes Elem hat eine Eltern-Phase, die nur eine Phase sein kann ...
                    Dim parentPhase As clsPhase = hproj.getParentPhaseByID(elemID)
                    If Not IsNothing(parentPhase) Then
                        allowedLeftDate = parentPhase.getStartDate
                        allowedRightDate = parentPhase.getEndDate
                    End If


                    Select Case currentColumn
                        ' Prüfuzng ob erlaubt notwendig 
                        Case col(PTmeTe.startdate)
                            If visboZustaende.currentZeileIsMilestone Then
                                ' bei Meilensteinen dürfte er gar nicht reinkommen ...
                                appInstance.Undo()
                            Else
                                ' erstmal prüfen, ob das neue Datum auch erlaubt ist, d.h innerhalb der zeitlichen Grenzen der Parent-Phase liegt 
                                ' und vor dem Ende-Datum der Phase liegt
                                Try
                                    Dim newStartDate As Date = CDate(Target.Value)
                                    If (newStartDate.Date >= allowedLeftDate.Date And newStartDate <= allowedRightDate) And newStartDate <= cphase.getEndDate Then
                                        ' alles ok, bearbeiten ..


                                    Else
                                        ' nicht ok, Datum liegt ausserhalb der erlaubten Grenzen 
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try

                            End If

                        ' Prüfung ob erlaubt notwendig  
                        Case col(PTmeTe.endDate)
                            ' erstmal prüfen, ob das neue Datum auch erlaubt ist, d.h innerhalb der zeitlichen Grenzen der Parent-Phase liegt 

                            If visboZustaende.currentZeileIsMilestone Then
                                Try
                                    Dim newEndeDate As Date = CDate(Target.Value)
                                    If (newEndeDate.Date >= allowedLeftDate.Date And newEndeDate <= allowedRightDate) Then
                                        ' alles ok, bearbeiten ..
                                        Call MsgBox("ok")

                                    Else
                                        ' nicht ok, Datum liegt ausserhalb der erlaubten Grenzen 
                                        Call MsgBox("nicht ok")
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try
                            Else
                                Try
                                    Dim newEndeDate As Date = CDate(Target.Value)
                                    If (newEndeDate.Date >= allowedLeftDate.Date And newEndeDate <= allowedRightDate) And newEndeDate >= cphase.getStartDate Then
                                        ' alles ok, bearbeiten ..
                                        Call MsgBox("ok")

                                    Else
                                        ' nicht ok, Datum liegt ausserhalb der erlaubten Grenzen 
                                        Call MsgBox("nicht ok")
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try
                            End If

                        ' Ampel-Status, inkl Prüfung
                        Case col(PTmeTe.trafficLight)

                            If Not IsNothing(Target.Value) Then
                                If IsNumeric(Target.Value) Then
                                    If CInt(Target.Value) >= 0 And CInt(Target.Value) <= 3 Then

                                        Dim colValue As Integer = CInt(Target.Value)
                                        Select Case colValue
                                            Case 0
                                                Target.Interior.Color = visboFarbeNone
                                            Case 1
                                                Target.Interior.Color = visboFarbeGreen
                                            Case 2
                                                Target.Interior.Color = visboFarbeYellow
                                            Case 3
                                                Target.Interior.Color = visboFarbeRed
                                        End Select

                                        If visboZustaende.currentZeileIsMilestone Then
                                            cMilestone.ampelStatus = colValue
                                        Else
                                            cphase.ampelStatus = colValue
                                        End If

                                    Else
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Else
                                    Target.Value = visboZustaende.oldValue
                                End If
                            Else
                                Target.Value = visboZustaende.oldValue
                            End If

                        ' Ampel-Erläuterung , alles als String erlaubt 
                        Case col(PTmeTe.explanation)

                            Dim myValue As String = ""
                            If Not IsNothing(Target.Value) Then
                                myValue = CStr(Target.Value)
                            End If

                            If visboZustaende.currentZeileIsMilestone Then
                                cMilestone.ampelErlaeuterung = myValue
                            Else
                                cphase.ampelErlaeuterung = myValue
                            End If


                        ' Verantwortlichkeit, später prüfen, ob als User existent 
                        Case col(PTmeTe.responsible)

                            Dim myValue As String = ""
                            If Not IsNothing(Target.Value) Then
                                myValue = CStr(Target.Value)
                            End If

                            If visboZustaende.currentZeileIsMilestone Then
                                cMilestone.verantwortlich = myValue
                            Else
                                cphase.verantwortlich = myValue
                            End If

                        Case col(PTmeTe.deliverables)

                            Dim myValue As String = ""

                            If Not IsNothing(Target.Value) Then
                                myValue = CStr(Target.Value)
                            End If

                            Dim tmpStr() As String = myValue.Split(New Char() {CChar(vbLf), CChar(vbCr)})

                            If visboZustaende.currentZeileIsMilestone Then
                                cMilestone.clearDeliverables()
                                For i As Integer = 0 To tmpStr.Length - 1
                                    cMilestone.addDeliverable(tmpStr(i))
                                Next

                            Else
                                cphase.clearDeliverables()
                                For i As Integer = 0 To tmpStr.Length - 1
                                    cphase.addDeliverable(tmpStr(i))
                                Next
                            End If


                        Case col(PTmeTe.percentDone)

                            Dim myValue As Double = 0.0

                            If Not IsNothing(Target.Value) Then
                                If IsNumeric(Target.Value) Then
                                    If CDbl(Target.Value) >= 0 And CDbl(Target.Value) <= 1.0 Then
                                        myValue = CDbl(Target.Value)

                                        If visboZustaende.currentZeileIsMilestone Then
                                            cMilestone.percentDone = myValue
                                        Else
                                            cphase.percentDone = myValue
                                        End If
                                    Else
                                        Target.Value = visboZustaende.oldValue
                                    End If
                                Else
                                    Target.Value = visboZustaende.oldValue
                                End If
                            Else
                                Target.Value = visboZustaende.oldValue
                            End If



                        Case col(PTmeTe.documentLink)

                            Dim myValue As String = ""

                            If Not IsNothing(Target.Value) Then
                                myValue = CStr(Target.Value).Trim
                            End If

                            If isValidURL(myValue) Or myValue = "" Then
                                If visboZustaende.currentZeileIsMilestone Then
                                    cMilestone.DocURL = myValue
                                Else
                                    cphase.DocURL = myValue
                                End If
                            Else
                                Target.Value = visboZustaende.oldValue
                            End If


                        Case Else
                            ' nichs tun , nicht erlaubt ..
                    End Select

                Else
                    ' es darf nur eine Zelle selektiert werden 
                    'appInstance.Undo()
                    'Call MsgBox("bitte nur eine Zelle selektieren ...")
                    appInstance.Undo()
                    'Target.Cells(1, 1).value = visboZustaende.oldValue

                End If
            Catch ex As Exception

            End Try

        End If

        appInstance.EnableEvents = True
    End Sub

    ''' <summary>
    ''' er kann hier eigentlich nur selektieren, was auch nicht gesperrt ist 
    ''' eine Unterscheidung zu enableSorting ist nicht notwendig  
    ''' </summary>
    ''' <param name="Target"></param>
    Private Sub Tabelle3_SelectionChange(Target As Range) Handles Me.SelectionChange

        appInstance.EnableEvents = False

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim elemNameID As String = ""
        Dim elemName As String = ""
        Dim zeileHasChanged As Boolean = False

        Dim oldElemNameID As String = visboZustaende.currentElemID


        Dim pname As String = ""
        Dim oldMsTaskName As String = ""

        Try
            ' wenn mehr wie eine Zelle selektiert wurde ...
            If Target.Cells.Count > 1 Then
                Target = CType(Target.Cells(1, 1), Excel.Range)
                Target.Select()
            End If

            ' kann ggf später ergänzt werden ... 
            If Target.Row <> visboZustaende.oldRow Then
                zeileHasChanged = True

                ' kann ggf später ergänzt werden ... 
                'Call SelectionMode(oldRow, False)
                'Call SelectionMode(Target.Row, True)
            End If

            ' alte Row merken 
            visboZustaende.oldRow = Target.Row

            oldColumn = Target.Column
            oldRow = Target.Row
            If Not IsNothing(CType(Target.Cells(1, 1), Excel.Range).Value) Then
                visboZustaende.oldValue = CStr(CType(Target.Cells(1, 1), Excel.Range).Value)
            Else
                visboZustaende.oldValue = ""
            End If
            CType(Target.Cells(1, 1), Excel.Range).Select()

        Catch ex As Exception
            Call MsgBox("Fehler bei Selection Change, Massen-Edit" & vbLf & ex.Message)
            appInstance.EnableEvents = True
        End Try

        ' in oldRow muss jetzt der entsprechende Projekt-Name und Phasen-Name ausgelesen werden .. 
        ' folgende Bedingung muss gesichert sein: alle Projekte, die in MassEdit aufgeführt sind , 
        ' sind sowohl in Showprojekte als auch in dbCacheProjekte

        Dim pNameChanged As Boolean = False
        Dim elemNameChanged As Boolean = False

        Dim isMilestone As Boolean = True
        If IsNothing(CType(appInstance.ActiveSheet.Cells(Target.Row, col(PTmeTe.startdate)), Excel.Range).Value) Then
            isMilestone = True
        Else
            isMilestone = CStr(CType(appInstance.ActiveSheet.Cells(Target.Row, col(PTmeTe.startdate)), Excel.Range).Value).Trim <> ""
        End If


        Dim curCell As Excel.Range = CType(appInstance.ActiveSheet.Cells(Target.Row, col(PTmeTe.pName)), Excel.Range)
        pname = CStr(curCell.Value)


        curCell = CType(appInstance.ActiveSheet.Cells(Target.Row, col(PTmeTe.elemName)), Excel.Range)
        If Not IsNothing(curCell.Comment) Then
            elemNameID = curCell.Comment.Text.Trim
            If elemNameID = "" Then
                Call calcHryElemKey(CStr(curCell.Value), isMilestone)
            End If
        Else
            Call calcHryElemKey(CStr(curCell.Value), isMilestone)
        End If

        isMilestone = elemIDIstMeilenstein(elemNameID)

        elemNameChanged = (elemNameID <> visboZustaende.currentElemID)
        visboZustaende.currentElemID = elemNameID

        If IsNothing(visboZustaende.currentProject) Then
            ' es wurde bisher kein lastProject geladen 
            If ShowProjekte.contains(pname) Then
                visboZustaende.currentProject = ShowProjekte.getProject(pname)
                visboZustaende.currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pname, visboZustaende.currentProject.variantName))
                pNameChanged = True
            End If

        ElseIf pname <> visboZustaende.currentProject.name Then
            ' muss neu geholt werden 
            If ShowProjekte.contains(pname) Then
                visboZustaende.currentProject = ShowProjekte.getProject(pname)
                visboZustaende.currentProjectinSession = sessionCacheProjekte.getProject(calcProjektKey(pname, visboZustaende.currentProject.variantName))
                pNameChanged = True
            End If
        End If

        ' jetzt muss die Phase bzw der Meilenstein aktualisiert werden 
        ' das wird implizit in der Klasse clsVisboZustaende gemacht: Methode getcurrentPhase oder get CurrentMilestone
        ' bzw currentzeileIsMilestone


        ' wenn pNameChanged und das Info-Fenster angezeigt wird, dann aktualisieren 
        Dim alreadyDone As Boolean = False

        ' das Projekt- und Portfolio Chart Zeichnen kommt erst noch ... 
        ' tk 13.12.19
        ''If pNameChanged Or elemNameChanged Then
        ''    ' aktualisieren der Projekt-Charts 
        ''    Call aktualisiereCharts(visboZustaende.lastProject, True, calledFromMassEdit:=True, currentRoleName:="")

        ''End If


        '' hier wird jetzt ggf das Role/Cost Portfolio Chart aktualisiert ..
        ''If Not IsNothing(elemNameID) Then
        ''    If "" <> rcName Then
        ''        If rcName <> "" And Not alreadyDone Then
        ''            selectedProjekte.Clear(False)
        ''            selectedProjekte.Add(visboZustaende.lastProject, False)
        ''            Call awinNeuZeichnenDiagramme(typus:=8, roleCost:=rcName)
        ''        End If
        ''    End If
        ''End If


        appInstance.EnableEvents = True

    End Sub

    Private Sub Tabelle3_Deactivate() Handles Me.Deactivate

        appInstance.ActiveWindow.SplitColumn = 0
        appInstance.ActiveWindow.SplitRow = 0

        Application.DisplayFormulaBar = False

    End Sub
End Class
