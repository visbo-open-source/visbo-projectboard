
Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop.Excel

Public Class Tabelle4

    Private columnStartData As Integer = 5
    Private columnEndData As Integer = 11
    Private columnDsc As Integer = 6
    Private oldColumn As Integer = 5
    Private oldRow As Integer = 2
    Private columnName As Integer = 2

    Private Sub Tabelle4_ActivateEvent() Handles Me.ActivateEvent
        ' in der Mass-Edit Termine sollen Header und Formular-Bar immer erhalten bleiben ...
        Application.DisplayFormulaBar = True

        'Dim filterRange As Excel.Range
        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meAT)), Excel.Worksheet)


        ' jetzt den Schutz aufheben , falls einer definiert ist 
        If meWS.ProtectContents Then
            meWS.Unprotect(Password:="x")
        End If

        Try
            ' die Anzahl maximaler Zeilen bestimmen 
            With visboZustaende
                .meMaxZeile = CType(meWS, Excel.Worksheet).UsedRange.Rows.Count
                .meColRC = 5
                .meColSD = 6
                .meColED = 10 + customFieldDefinitions.count
                .meColpName = 2
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
                'CType(meWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
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
                             AllowInsertingRows:=False,
                             AllowDeletingColumns:=False,
                             AllowDeletingRows:=False,
                             AllowSorting:=True,
                             AllowFiltering:=True)
                    .EnableSelection = XlEnableSelection.xlUnlockedCells
                    .EnableAutoFilter = True
                End With
            End If


        Catch ex As Exception

        End Try


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
            Dim bedingung As Boolean = CBool(CType(meWS.Cells(cz, columnDsc + 1), Excel.Range).Locked = True) And Not eof

            Do While bedingung
                cz = cz + 1
                eof = (cz > visboZustaende.meMaxZeile)
                bedingung = CBool(CType(meWS.Cells(cz, columnDsc + 1), Excel.Range).Locked = True) And Not eof
            Loop

            If Not eof Then
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnDsc + 1), Excel.Range).Select()

                Dim pName As String = ""

                With visboZustaende

                    pName = CStr(CType(meWS.Cells(cz, visboZustaende.meColpName), Excel.Range).Value)
                    If ShowProjekte.contains(pName) Then
                        .lastProject = ShowProjekte.getProject(pName)
                        .lastProjectDB = dbCacheProjekte.getProject(calcProjektKey(pName, .lastProject.variantName))
                    End If

                End With
            Else
                CType(CType(meWS, Excel.Worksheet).Cells(cz, columnDsc), Excel.Range).Locked = False
            End If

            CType(CType(meWS, Excel.Worksheet).Cells(cz, columnDsc), Excel.Range).Select()

        Catch ex As Exception

        End Try

        Application.EnableEvents = formerEE
        If Application.ScreenUpdating = False Then
            Application.ScreenUpdating = True
        End If


    End Sub

    Private Sub Tabelle4_Deactivate() Handles Me.Deactivate

        appInstance.ActiveWindow.SplitColumn = 0
        appInstance.ActiveWindow.SplitRow = 0
        appInstance.DisplayFormulaBar = False

    End Sub
End Class
