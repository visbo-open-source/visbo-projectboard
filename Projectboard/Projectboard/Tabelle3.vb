
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Microsoft.Office.Interop.Excel
''' <summary>
''' 
''' </summary>
Public Class Tabelle3

    Private columnStartData As Integer = 5
    Private columnEndData As Integer = 11
    Private columnTE As Integer = 5
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private columnName As Integer = 2

    Private Sub Tabelle3_ActivateEvent() Handles Me.ActivateEvent

        ' in der Mass-Edit Termine sollen Header und Formular-Bar immer erhalten bleiben ...
        'Application.DisplayFormulaBar = False

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
                .meMaxZeile = CType(appInstance.ActiveSheet, Excel.Worksheet).UsedRange.Rows.Count
                .meColRC = 5
                .meColSD = 5
                .meColED = 11
                .meColpName = 2
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


            'With meWS
            '    filterRange = CType(.Range(.Cells(1, 1), .Cells(1, 11)), Excel.Range)
            'End With

            ' jetzt die Autofilter aktivieren ... 
            If Not CType(meWS, Excel.Worksheet).AutoFilterMode = True Then
                'CType(meWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                CType(meWS, Excel.Worksheet).Rows(1).AutoFilter()

                '' jetzt überprüfen, ob nur eine bestimmte Rolle/Kostenart angezeigt, d.h gefiltert werden soll  
                'If Not IsNothing(rcName) Then
                '    CType(CType(meWS, Excel.Worksheet).Rows(1), Excel.Range).AutoFilter(Field:=visboZustaende.meColRC, Criteria1:=rcName)
                'End If

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

        With meWS
            CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
        End With

        If Not IsNothing(appInstance.ActiveCell) Then
            visboZustaende.oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If


        ' einen Select machen - nachdem Event Behandlung wieder true ist, dann werden project und lastprojectDB gesetzt ...
        Try
            'CType(CType(meWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
            ' jetzt auf die erste selektierbare Zeile gehen ... 
            Dim cz As Integer = 2
            Dim eof As Boolean = (cz > visboZustaende.meMaxZeile)
            ' auf TE+1 = End-Datum platzieren, weil das immer zu editieren ist 
            Dim bedingung As Boolean = CBool(CType(meWS.Cells(cz, columnTE + 1), Excel.Range).Locked = True) And Not eof

            Do While bedingung
                cz = cz + 1
                eof = (cz > visboZustaende.meMaxZeile)
                bedingung = CBool(CType(meWS.Cells(cz, columnTE + 1), Excel.Range).Locked = True) And Not eof
            Loop

            If Not eof Then
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnTE + 1), Excel.Range).Select()

                Dim pName As String = ""

                With visboZustaende

                    pName = CStr(CType(appInstance.ActiveSheet.Cells(cz, visboZustaende.meColpName), Excel.Range).Value)
                    If ShowProjekte.contains(pName) Then
                        .lastProject = ShowProjekte.getProject(pName)
                        .lastProjectDB = dbCacheProjekte.getProject(calcProjektKey(pName, .lastProject.variantName))
                    End If

                End With
            Else
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnTE), Excel.Range).Locked = False
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnTE), Excel.Range).Select()
            End If

        Catch ex As Exception

        End Try

        Application.EnableEvents = formerEE
        If Application.ScreenUpdating = False Then
            Application.ScreenUpdating = True
        End If


    End Sub
End Class
