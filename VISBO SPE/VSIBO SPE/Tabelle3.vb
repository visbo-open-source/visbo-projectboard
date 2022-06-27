
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
        invoiceValue = 12
        invoiceTerm = 13
        penaltyValue = 14
        penaltyDate = 15
    End Enum

    ' enthält die Spalten, wo die einzelnen Felder stehen , korreliert mit der Enum allianzSpalten
    Private col() As Integer

    Private Sub Tabelle3_ActivateEvent() Handles Me.ActivateEvent

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        ' in der Mass-Edit Termine sollen Header und Formular-Bar immer erhalten bleiben ...
        Try
            Application.DisplayFormulaBar = False
        Catch ex As Exception

        End Try

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
        col(PTmeTe.invoiceValue) = 13
        col(PTmeTe.invoiceTerm) = 14
        col(PTmeTe.penaltyValue) = 15
        col(PTmeTe.penaltyDate) = 16

        ' initial setzen der Spalten ... 

        'Dim filterRange As Excel.Range

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

        '' jetzt die Splaten für ProjNr, ProjName, VariantenName ausblenden

        ''?????
        appInstance.EnableEvents = True
        Dim aa As Boolean = appInstance.EnableEvents

        ' jetzt die Spalte 6 einblenden bzw. ausblenden 
        Try
            'If visboZustaende.projectBoardMode = ptModus.massEditTermine Then
            'CType(meWS.Columns(6), Excel.Range).EntireColumn.Hidden = True
            If ShowProjekte.Count = 1 Then
                CType(meWS.Columns("A"), Excel.Range).Hidden = True
                CType(meWS.Columns("B"), Excel.Range).Hidden = True
                CType(meWS.Columns("C"), Excel.Range).Hidden = True
            Else
                CType(meWS.Columns("A"), Excel.Range).Hidden = False
                CType(meWS.Columns("B"), Excel.Range).Hidden = False
                CType(meWS.Columns("c"), Excel.Range).Hidden = False
            End If
            'ElseIf visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
            '    If RoleDefinitions.getAllSkillIDs.Count > 0 Then
            '        CType(meWS.Columns(6), Excel.Range).EntireColumn.Hidden = False
            '    Else
            '        CType(meWS.Columns(6), Excel.Range).EntireColumn.Hidden = True
            '    End If
            '    If ShowProjekte.Count = 1 Then
            '        CType(meWS.Columns(1), Excel.Range).EntireColumn.Hidden = True
            '        CType(meWS.Columns(2), Excel.Range).EntireColumn.Hidden = True
            '        CType(meWS.Columns(3), Excel.Range).EntireColumn.Hidden = True
            '    Else
            '        CType(meWS.Columns(1), Excel.Range).EntireColumn.Hidden = False
            '        CType(meWS.Columns(2), Excel.Range).EntireColumn.Hidden = False
            '        CType(meWS.Columns(3), Excel.Range).EntireColumn.Hidden = False
            '    End If

            'End If
        Catch ex As Exception

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
                .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
                .EnableAutoFilter = True
            End With


        Catch ex As Exception

        End Try


        appInstance.EnableEvents = formerEE

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

            If massColFontValues(1, 0) <> 0 Then
                .Zoom = massColFontValues(1, 0)
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

        '' jetzt die Splaten für ProjNr, ProjName, VariantenName ausblenden
        'If ShowProjekte.Count = 1 Then
        '    CType(meWS.Columns(1), Excel.Range).EntireColumn.Hidden = True
        '    CType(meWS.Columns(2), Excel.Range).EntireColumn.Hidden = True
        '    CType(meWS.Columns(3), Excel.Range).EntireColumn.Hidden = True
        'Else
        '    CType(meWS.Columns(1), Excel.Range).EntireColumn.Hidden = False
        '    CType(meWS.Columns(2), Excel.Range).EntireColumn.Hidden = False
        '    CType(meWS.Columns(3), Excel.Range).EntireColumn.Hidden = False
        'End If

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

            Dim oldMovableStatus As Boolean = hproj.movable
            Dim oldProjektStatus As String = hproj.vpStatus
            Dim oldVariantName As String = hproj.variantName

            hproj.variantName = "$tmpv1"
            'ur: 211202: hproj.Status = ProjektStatus(PTProjektStati.geplant)
            hproj.vpStatus = VProjectStatus(PTVPStati.initialized)
            hproj.movable = True

            Dim allowedLeftDate As Date = StartofCalendar
            If hproj.hasActualValues Then
                allowedLeftDate = getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, False)
            End If
            Dim allowedRightDate As Date = StartofCalendar.AddYears(20).AddDays(-1)

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
                        If hproj.hasActualValues Then
                            If parentPhase.getStartDate < hproj.actualDataUntil Then
                                allowedLeftDate = getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, False)
                            End If
                        End If

                        allowedRightDate = parentPhase.getEndDate

                    End If


                    Select Case currentColumn
                        ' Prüfung ob erlaubt notwendig 

                        Case col(PTmeTe.startdate)

                            ' hier kann es nur eine Phase sein 
                            ' das ggf eingegebene Datum wird geprüft und das Formular aufgeschaltet ... 

                            Try
                                Dim newStartDate As Date = CDate(Target.Value)
                                If (newStartDate.Date >= allowedLeftDate.Date And newStartDate <= allowedRightDate) And newStartDate <= cphase.getEndDate Then
                                    ' alles ok, bearbeiten ..

                                    ' jetzt muss der neue Offset in Tagen bestimmt werden ... 
                                    Dim newOffsetInTagen As Long = DateDiff(DateInterval.Day, hproj.startDate.Date, newStartDate.Date)
                                    Dim newDauerInTagen As Long = DateDiff(DateInterval.Day, newStartDate, cphase.getEndDate) + 1
                                    Dim autoAdjustChilds As Boolean = True


                                    If cphase.nameID = rootPhaseName Then

                                        hproj.startDate = newStartDate
                                        newOffsetInTagen = 0

                                    End If

                                    ' jetzt wird die Phase entsprechend geändert ...
                                    ' jetzt kommt der rekursive Aufruf: die Phase mit all ihren Kindern und Kindeskindern wird angepasst
                                    ' unter Berücksichtigung der Ist-Daten, falls welche existieren ...  

                                    Dim nameIDCollection As Collection = hproj.getAllChildIDsOf(elemID)
                                    cphase = cphase.adjustPhaseAndChilds(newOffsetInTagen, newDauerInTagen, autoAdjustChilds)

                                    ' tk 4.1.20 eigentlich braucht man das hier nicht mehr ... 
                                    'Dim diffDays As Long = DateDiff(DateInterval.Day, hproj.startDate.Date, newStartDate.Date)
                                    'If diffDays <> 0 Then
                                    '    ' tk 30.12.19 hier muss sichergestellt sein, dass die 
                                    '    Call hproj.syncXWertePhases()
                                    'End If

                                    ' jetzt werden die Excel Zeilen aktualisiert 
                                    If autoAdjustChilds And nameIDCollection.Count > 0 Then
                                        ' 
                                        Try
                                            Dim currentChildRow As Integer = Target.Row + 1
                                            Dim potentialChildID As String = CStr(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment.text)
                                            Dim isChild As Boolean = nameIDCollection.Contains(potentialChildID)


                                            Do While isChild
                                                Dim isMilestone As Boolean = elemIDIstMeilenstein(potentialChildID)
                                                If isMilestone Then
                                                    Dim tmpMS As clsMeilenstein = hproj.getMilestoneByID(potentialChildID)
                                                    meWS.Cells(currentChildRow, col(PTmeTe.startdate)).value = ""
                                                    meWS.Cells(currentChildRow, col(PTmeTe.endDate)).value = tmpMS.getDate
                                                Else
                                                    Dim tmpPh As clsPhase = hproj.getPhaseByID(potentialChildID)
                                                    meWS.Cells(currentChildRow, col(PTmeTe.startdate)).value = tmpPh.getStartDate
                                                    meWS.Cells(currentChildRow, col(PTmeTe.endDate)).value = tmpPh.getEndDate
                                                End If

                                                currentChildRow = currentChildRow + 1

                                                Try
                                                    If Not IsNothing(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment) Then
                                                        potentialChildID = CStr(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment.text)
                                                        If potentialChildID <> "" Then
                                                            isChild = nameIDCollection.Contains(potentialChildID)
                                                        Else
                                                            isChild = False
                                                        End If
                                                    Else
                                                        isChild = False
                                                    End If
                                                Catch ex As Exception
                                                    isChild = False
                                                End Try


                                            Loop
                                        Catch ex As Exception

                                        End Try


                                    End If


                                Else
                                    ' nicht ok, Datum liegt ausserhalb der erlaubten Grenzen 
                                    Target.Value = visboZustaende.oldValue
                                End If


                            Catch ex As Exception
                                Target.Value = visboZustaende.oldValue
                            End Try


                        ' Prüfung ob erlaubt notwendig  
                        Case col(PTmeTe.endDate)

                            ' hier kann es eine Phase oder ein Meilenstein sein ... 

                            If visboZustaende.currentZeileIsMilestone Then
                                ' Meilenstein 
                                Try
                                    Dim newEndDate As Date = CDate(Target.Value)
                                    If (newEndDate >= allowedLeftDate.Date And newEndDate <= allowedRightDate) Then
                                        ' alles ok, bearbeiten ..
                                        cMilestone.setDate = newEndDate
                                    Else
                                        ' nicht ok, Datum liegt ausserhalb der erlaubten Grenzen 
                                        Target.Value = visboZustaende.oldValue
                                    End If


                                Catch ex As Exception
                                    Target.Value = visboZustaende.oldValue
                                End Try


                            Else
                                ' Phase 
                                Try
                                    Dim newEndDate As Date = CDate(Target.Value)
                                    If cphase.nameID = rootPhaseName Then
                                        ' das kleinste zugelassene Datum ist das Ende des Monats , der dem ActualDataUntil folgt ...
                                        If hproj.hasActualValues Then
                                            allowedLeftDate = getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, True)
                                        End If

                                    End If

                                    If (newEndDate.Date >= allowedLeftDate.Date And newEndDate <= allowedRightDate) And newEndDate >= cphase.getStartDate Then
                                        ' alles ok, bearbeiten ..

                                        ' jetzt muss die neue Dauer in Tagen bestimmt werden ... 
                                        Dim newDauerInTagen As Long = DateDiff(DateInterval.Day, cphase.getStartDate, newEndDate) + 1
                                        Dim newOffsetInTagen As Long = cphase.startOffsetinDays

                                        ' jetzt wird die Phase entsprechend geändert ...
                                        ' jetzt kommt der rekursive Aufruf: die Phase mit all ihren Kindern und Kindeskindern wird angepasst
                                        ' unter Berücksichtigung der Ist-Daten, falls welche existieren ...  
                                        Dim autoAdjustChilds As Boolean = True
                                        Dim nameIDCollection As Collection = hproj.getAllChildIDsOf(elemID)
                                        cphase = cphase.adjustPhaseAndChilds(newOffsetInTagen, newDauerInTagen, autoAdjustChilds)

                                        ' braucht man das hier ...? 
                                        'Call hproj.syncXWertePhases()

                                        ' jetzt die Excel Zeilen der Kinder aktualisieren  
                                        If autoAdjustChilds And nameIDCollection.Count > 0 Then
                                            ' 
                                            Try
                                                Dim currentChildRow As Integer = Target.Row + 1
                                                Dim potentialChildID As String = CStr(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment.text)
                                                Dim isChild As Boolean = nameIDCollection.Contains(potentialChildID)


                                                Do While isChild
                                                    Dim isMilestone As Boolean = elemIDIstMeilenstein(potentialChildID)
                                                    If isMilestone Then
                                                        Dim tmpMS As clsMeilenstein = hproj.getMilestoneByID(potentialChildID)
                                                        meWS.Cells(currentChildRow, col(PTmeTe.startdate)).value = ""
                                                        meWS.Cells(currentChildRow, col(PTmeTe.endDate)).value = tmpMS.getDate
                                                    Else
                                                        Dim tmpPh As clsPhase = hproj.getPhaseByID(potentialChildID)
                                                        meWS.Cells(currentChildRow, col(PTmeTe.startdate)).value = tmpPh.getStartDate
                                                        meWS.Cells(currentChildRow, col(PTmeTe.endDate)).value = tmpPh.getEndDate
                                                    End If

                                                    currentChildRow = currentChildRow + 1

                                                    Try
                                                        If Not IsNothing(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment) Then
                                                            potentialChildID = CStr(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment.text)
                                                            If potentialChildID <> "" Then
                                                                isChild = nameIDCollection.Contains(potentialChildID)
                                                            Else
                                                                isChild = False
                                                            End If
                                                        Else
                                                            isChild = False
                                                        End If
                                                    Catch ex As Exception
                                                        isChild = False
                                                    End Try


                                                Loop
                                            Catch ex As Exception

                                            End Try


                                        End If


                                    Else
                                        ' nicht ok, Datum liegt ausserhalb der erlaubten Grenzen 
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
                                                'Target.Interior.Color = visboFarbeNone
                                                'Target.Interior.ColorIndex = -4142
                                                Target.Interior.ColorIndex = XlColorIndex.xlColorIndexNone
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

                        Case col(PTmeTe.invoiceValue)
                            Dim myValue As Double = 0.0

                            If Not IsNothing(Target.Value) Then
                                If IsNumeric(Target.Value) Then
                                    If CDbl(Target.Value) >= 0 Then
                                        myValue = CDbl(Target.Value)

                                        If visboZustaende.currentZeileIsMilestone Then
                                            Dim newInvoice As New KeyValuePair(Of Double, Integer)(myValue, cMilestone.invoice.Value)
                                            If myValue = 0 Then
                                                newInvoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
                                                ' Terms of payment anpassen 
                                                meWS.Cells(currentZeile, currentColumn + 1).value = ""
                                            End If
                                            cMilestone.invoice = newInvoice
                                        Else
                                            Dim newInvoice As New KeyValuePair(Of Double, Integer)(myValue, cphase.invoice.Value)
                                            If myValue = 0 Then
                                                newInvoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
                                                ' Terms of payment anpassen 
                                                meWS.Cells(currentZeile, currentColumn + 1).value = ""
                                            End If
                                            cphase.invoice = newInvoice
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

                        Case col(PTmeTe.invoiceTerm)
                            Dim myValue As Integer = 0

                            If Not IsNothing(Target.Value) Then
                                If IsNumeric(Target.Value) Then
                                    If CInt(Target.Value) > 0 Then
                                        myValue = CInt(Target.Value)

                                        If visboZustaende.currentZeileIsMilestone Then
                                            Dim newInvoice As New KeyValuePair(Of Double, Integer)(cMilestone.invoice.Key, myValue)
                                            cMilestone.invoice = newInvoice
                                        Else
                                            Dim newInvoice As New KeyValuePair(Of Double, Integer)(cphase.invoice.Key, myValue)
                                            cphase.invoice = newInvoice
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

                        Case col(PTmeTe.penaltyDate)
                            Dim myValue As Date = Date.MinValue

                            If Not IsNothing(Target.Value) Then
                                If IsDate(Target.Value) Then
                                    If CDate(Target.Value) > hproj.startDate Then
                                        myValue = CDate(Target.Value)

                                        If visboZustaende.currentZeileIsMilestone Then
                                            Dim newPenalty As New KeyValuePair(Of Date, Double)(myValue, cMilestone.penalty.Value)
                                            cMilestone.penalty = newPenalty
                                        Else
                                            Dim newPenalty As New KeyValuePair(Of Date, Double)(myValue, cphase.penalty.Value)
                                            cphase.penalty = newPenalty
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

                        Case col(PTmeTe.penaltyValue)
                            Dim myValue As Double = 0.0

                            If Not IsNothing(Target.Value) Then
                                If IsNumeric(Target.Value) Then
                                    If CDbl(Target.Value) >= 0 Then
                                        myValue = CDbl(Target.Value)

                                        If visboZustaende.currentZeileIsMilestone Then
                                            Dim newPenalty As New KeyValuePair(Of Date, Double)(cMilestone.penalty.Key, myValue)
                                            If myValue = 0 Then
                                                newPenalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)
                                                meWS.Cells(currentZeile, currentColumn + 1).value = ""
                                            End If
                                            cMilestone.penalty = newPenalty
                                        Else
                                            Dim newPenalty As New KeyValuePair(Of Date, Double)(cphase.penalty.Key, myValue)
                                            If myValue = 0 Then
                                                newPenalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)
                                                meWS.Cells(currentZeile, currentColumn + 1).value = ""
                                            End If
                                            cphase.penalty = newPenalty
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

                        Case Else
                            ' nichs tun , nicht erlaubt ..
                    End Select

                    If Not IsNothing(Target.Cells(1, 1).value) Then
                        visboZustaende.oldValue = CStr(Target.Cells(1, 1).value)
                    Else
                        visboZustaende.oldValue = ""
                    End If

                Else
                    ' es darf nur eine Zelle selektiert werden 
                    'appInstance.Undo()
                    'Call MsgBox("bitte nur eine Zelle selektieren ...")
                    appInstance.Undo()
                    'Target.Cells(1, 1).value = visboZustaende.oldValue

                End If
            Catch ex As Exception

            End Try

            ' jetzt wieder zurücksetzen 

            hproj.movable = oldMovableStatus
            hproj.vpStatus = oldProjektStatus
            hproj.variantName = oldVariantName


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
                elemNameID = calcHryElemKey(CStr(curCell.Value), isMilestone)
            End If
        Else
            elemNameID = calcHryElemKey(CStr(curCell.Value), isMilestone)
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

        If ShowProjekte.Count = 1 Then
            CType(meWS.Columns(1), Excel.Range).EntireColumn.Hidden = True
            CType(meWS.Columns(2), Excel.Range).EntireColumn.Hidden = True
            CType(meWS.Columns(3), Excel.Range).EntireColumn.Hidden = True
        Else
            CType(meWS.Columns(1), Excel.Range).EntireColumn.Hidden = False
            CType(meWS.Columns(2), Excel.Range).EntireColumn.Hidden = False
            CType(meWS.Columns(3), Excel.Range).EntireColumn.Hidden = False
        End If

    End Sub

    Private Sub Tabelle3_Deactivate() Handles Me.Deactivate

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
            'CType(Me.Columns(1), Excel.Range).EntireColumn.Hidden = False
            'CType(Me.Columns(2), Excel.Range).EntireColumn.Hidden = False
            'CType(Me.Columns(3), Excel.Range).EntireColumn.Hidden = False
            'CType(Me.Columns(6), Excel.Range).EntireColumn.Hidden = False
        Catch ex As Exception

        End Try

        Try
            appInstance.DisplayFormulaBar = False
        Catch ex As Exception

        End Try


    End Sub



    Private Sub Tabelle3_BeforeRightClick(Target As Range, ByRef Cancel As Boolean) Handles Me.BeforeRightClick

        Dim hproj As clsProjekt = visboZustaende.currentProject
        Dim cphase As clsPhase = Nothing
        Dim cMilestone As clsMeilenstein = Nothing

        Dim oldMovableStatus As Boolean = hproj.movable
        Dim oldProjektStatus As String = hproj.vpStatus
        Dim oldVariantName As String = hproj.variantName

        hproj.variantName = "$tmpv1"
        'ur: 211202: hproj.Status = ProjektStatus(PTProjektStati.geplant)
        hproj.vpStatus = VProjectStatus(PTVPStati.initialized)
        hproj.movable = True

        'Dim allowedLeftDate As Date = hproj.startDate
        'Dim allowedRightDate As Date = hproj.endeDate

        Dim allowedLeftDate As Date = StartofCalendar
        If hproj.hasActualValues Then
            allowedLeftDate = getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, False)
        End If
        Dim allowedRightDate As Date = StartofCalendar.AddYears(20).AddDays(-1)

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)

        appInstance.EnableEvents = False
        Try
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

                    If hproj.hasActualValues Then
                        If parentPhase.getStartDate < hproj.actualDataUntil Then
                            allowedLeftDate = getDateofColumn(getColumnOfDate(hproj.actualDataUntil) + 1, False)
                        End If
                    End If

                    allowedRightDate = parentPhase.getEndDate
                End If


                If Target.Column = col(PTmeTe.startdate) Or Target.Column = col(PTmeTe.endDate) Then

                    If visboZustaende.currentZeileIsMilestone Then

                        ' Meilenstein

                        ' in target.Value ist jetzt der neue Wert
                        Dim frmDateEdit As New frmEditDates

                        frmDateEdit.lblElemName.Text = elemNameOfElemID(visboZustaende.currentElemID)
                        frmDateEdit.startdatePicker.Value = cMilestone.getDate
                        frmDateEdit.startdatePicker.Enabled = False

                        ' Checkbox Auto Distribution is invisible ..
                        frmDateEdit.chkbxAutoDistr.Visible = False

                        frmDateEdit.enddatePicker.Value = cMilestone.getDate
                        frmDateEdit.IsMilestone = True

                        frmDateEdit.allowedDateLeft = allowedLeftDate
                        frmDateEdit.allowedDateRight = allowedRightDate

                        If frmDateEdit.ShowDialog() = DialogResult.OK Then
                            Target.Value = frmDateEdit.enddatePicker.Value
                            cMilestone.setDate = CDate(Target.Value)
                        Else
                            Target.Value = visboZustaende.oldValue
                        End If


                    Else
                        ' ist Phase ...

                        Dim frmDateEdit As New frmEditDates
                        Dim wasRootPhase As Boolean = False

                        ' wenn die Phase Kinder hat, muss das Flag "automatisch anpassen" angezeigt werden 
                        Dim anzChilds As Integer = hproj.hierarchy.getChildIDsOf(cphase.nameID, True).Count + hproj.hierarchy.getChildIDsOf(cphase.nameID, False).Count
                        If anzChilds > 0 Then
                            frmDateEdit.chkbx_adjustChilds.Visible = False
                            frmDateEdit.chkbx_adjustChilds.Enabled = False
                            frmDateEdit.chkbx_adjustChilds.Checked = awinSettings.autoAjustChilds
                        End If

                        ' Checkbox Auto Distribution is visible ..
                        frmDateEdit.chkbxAutoDistr.Visible = False
                        frmDateEdit.chkbxAutoDistr.Checked = Not awinSettings.noNewCalculation

                        frmDateEdit.lblElemName.Text = elemNameOfElemID(visboZustaende.currentElemID)
                        frmDateEdit.IsMilestone = False

                        frmDateEdit.startdatePicker.Value = cphase.getStartDate

                        If allowedLeftDate > cphase.getStartDate Then
                            frmDateEdit.startdatePicker.Enabled = False
                        End If

                        frmDateEdit.enddatePicker.Value = cphase.getEndDate

                        frmDateEdit.allowedDateLeft = allowedLeftDate
                        frmDateEdit.allowedDateRight = allowedRightDate

                        If frmDateEdit.ShowDialog() = DialogResult.OK Then
                            ' jetzt muss der neue Offset in Tagen bestimmt werden ... 
                            ' hier ist bereits im Formular sichergestellt, dass es sich um valide Datum-Angaben handelt .. 
                            ' ur:20220609: hier nicht benötigt:::awinSettings.noNewCalculation = Not frmDateEdit.chkbxAutoDistr.Checked

                            Dim newOffsetInTagen As Long = DateDiff(DateInterval.Day, hproj.startDate.Date, frmDateEdit.startdatePicker.Value.Date)
                            Dim newDauerInTagen As Long = DateDiff(DateInterval.Day, frmDateEdit.startdatePicker.Value.Date, frmDateEdit.enddatePicker.Value.Date) + 1

                            'ur;09062022: wird ersetzt durch awinSetting.autoAjustChilds:Dim autoAdjustChilds As Boolean = frmDateEdit.chkbx_adjustChilds.Checked

                            If cphase.nameID = rootPhaseName Then

                                wasRootPhase = True

                                Dim diffDays As Long = DateDiff(DateInterval.Day, hproj.startDate.Date, frmDateEdit.startdatePicker.Value.Date)
                                hproj.startDate = frmDateEdit.startdatePicker.Value

                                If diffDays <> 0 Then
                                    ' tk 30.12.19 hier muss sichergestellt sein, dass die X-Werte neu berechnet werden, denn es kann sein, 
                                    ' dass so verschoben wird, dass offsets und Dauern jeweils gleich sind. 
                                    ' 
                                    Call hproj.syncXWertePhases()
                                End If

                                newOffsetInTagen = 0

                            End If

                            ' jetzt kommt der rekursive Aufruf: die Phase mit all ihren Kindern und Kindeskindern wird angepasst
                            ' unter Berücksichtigung der Ist-Daten, falls welche existieren ...  
                            Dim nameIDCollection As Collection = hproj.getAllChildIDsOf(elemID)
                            cphase = cphase.adjustPhaseAndChilds(newOffsetInTagen, newDauerInTagen, awinSettings.autoAjustChilds)

                            ' jetzt die Excel Zellen der aktuellen Zeile, der Phase anpassen ... 
                            meWS.Cells(Target.Row, col(PTmeTe.startdate)).value = frmDateEdit.startdatePicker.Value
                            meWS.Cells(Target.Row, col(PTmeTe.endDate)).value = frmDateEdit.enddatePicker.Value

                            If awinSettings.autoAjustChilds And nameIDCollection.Count > 0 Then

                                Try
                                    ' jetzt die Excel Zeilen der Kinder aktualisieren  
                                    Dim currentChildRow As Integer = Target.Row + 1
                                    Dim potentialChildID As String = CStr(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment.text)
                                    Dim isChild As Boolean = nameIDCollection.Contains(potentialChildID)


                                    Do While isChild
                                        Dim isMilestone As Boolean = elemIDIstMeilenstein(potentialChildID)
                                        If isMilestone Then
                                            Dim tmpMS As clsMeilenstein = hproj.getMilestoneByID(potentialChildID)
                                            meWS.Cells(currentChildRow, col(PTmeTe.startdate)).value = ""
                                            meWS.Cells(currentChildRow, col(PTmeTe.endDate)).value = tmpMS.getDate
                                        Else
                                            Dim tmpPh As clsPhase = hproj.getPhaseByID(potentialChildID)
                                            meWS.Cells(currentChildRow, col(PTmeTe.startdate)).value = tmpPh.getStartDate
                                            meWS.Cells(currentChildRow, col(PTmeTe.endDate)).value = tmpPh.getEndDate
                                        End If

                                        currentChildRow = currentChildRow + 1

                                        Try
                                            If Not IsNothing(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment) Then
                                                potentialChildID = CStr(meWS.Cells(currentChildRow, col(PTmeTe.elemName)).comment.text)
                                                If potentialChildID <> "" Then
                                                    isChild = nameIDCollection.Contains(potentialChildID)
                                                Else
                                                    isChild = False
                                                End If
                                            Else
                                                isChild = False
                                            End If
                                        Catch ex As Exception
                                            isChild = False
                                        End Try


                                    Loop

                                Catch ex As Exception

                                End Try

                            End If


                        Else
                            Target.Value = visboZustaende.oldValue
                        End If

                    End If

                Else
                    appInstance.EnableEvents = True
                    Cancel = True
                End If
            End If

        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try

        ' jetzt wieder zurücksetzen 

        hproj.movable = oldMovableStatus
        hproj.vpStatus = oldProjektStatus
        hproj.variantName = oldVariantName

        appInstance.EnableEvents = True
        Cancel = True
    End Sub

    Private Sub Tabelle3_Startup(sender As Object, e As EventArgs) Handles Me.Startup
        If visboClient.Contains("VISBO SPE") Then
            'Call MsgBox("bin im meTE")
        End If
    End Sub
End Class
