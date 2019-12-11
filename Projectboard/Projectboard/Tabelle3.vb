
Imports ProjectBoardDefinitions
Imports Microsoft.Office.Interop.Excel

Public Class Tabelle3

    Private columnStartDate As Integer = 5
    Private columnEndDate As Integer = 6

    Private oldColumn As Integer = 6
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
                         AllowSorting:=True,
                         AllowFiltering:=True)
                .EnableSelection = XlEnableSelection.xlUnlockedCells
                .EnableAutoFilter = True
            End With


        Catch ex As Exception

        End Try


        If Not IsNothing(appInstance.ActiveCell) Then
            visboZustaende.oldValue = CStr(CType(appInstance.ActiveCell, Excel.Range).Value)
        End If

        ' es wird erst mal kein automatischer Select gemacht ... 

        Application.EnableEvents = formerEE
        If Application.ScreenUpdating = False Then
            Application.ScreenUpdating = True
        End If


    End Sub

    Private Sub Tabelle3_Change(Target As Range) Handles Me.Change

        ' damit nicht eine immerwährende Event Orgie durch Änderung in den Zellen abgeht ...
        appInstance.EnableEvents = False

        Dim currentCell As Excel.Range = Target

        Try
            Dim datesWereChanged As Boolean = False

            Dim meWB As Excel.Workbook = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
            Dim meWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meTE)), Excel.Worksheet)


            If Target.Cells.Count = 1 Then

            Else
                ' es darf nur eine Zelle selektiert werden 
                'appInstance.Undo()
                'Call MsgBox("bitte nur eine Zelle selektieren ...")
                appInstance.Undo()
                'Target.Cells(1, 1).value = visboZustaende.oldValue

            End If
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True
    End Sub
End Class
