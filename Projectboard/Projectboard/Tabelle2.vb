
Imports ProjectBoardDefinitions

Public Class Tabelle2

    Private Sub Tabelle2_ActivateEvent() Handles Me.ActivateEvent

        Dim rng As Excel.Range
        Dim tmpstart As Date
        Application.DisplayFormulaBar = True

        Dim formerEE As Boolean = Application.EnableEvents
        Application.EnableEvents = False

        Application.ScreenUpdating = False

        ' bei betreten dieses Tabellenblattes soll es auf false gesetzt werden - 
        ' in dem Moment, wo tabelle1 wieder aktiviert wird, also bei tabelle1.activate wird es auf true gesetzt ... 

        enableOnUpdate = False


        With Application.ActiveSheet

            If awinSettings.zeitEinheit = "PM" Then

                .cells(1, 1).value = "Monate"

                rng = .Range(.cells(1, 3), .cells(1, 4))
                rng.NumberFormat = "mmm-yy"

                Dim destinationRange As Excel.Range = .Range(.Cells(1, 3), .Cells(1, 62))
                With destinationRange
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                    .NumberFormat = "mmm-yy"
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = Excel.Constants.xlContext
                    .MergeCells = False
                    .Interior.color = noshowtimezone_color
                End With

                rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

            ElseIf awinSettings.zeitEinheit = "PW" Then
                .cells(1, 1).value = "Wochen"
                For i = 1 To 210
                    CType(.cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddDays((i - 1) * 7)
                Next

            ElseIf awinSettings.zeitEinheit = "PT" Then
                .cells(1, 1).value = "Tage"
                Dim workOnSat As Boolean = False
                Dim workOnSun As Boolean = False


                If Weekday(StartofCalendar, FirstDayOfWeek.Monday) > 3 Then
                    tmpstart = StartofCalendar.AddDays(8 - Weekday(StartofCalendar, FirstDayOfWeek.Monday))
                Else
                    tmpstart = StartofCalendar.AddDays(Weekday(StartofCalendar, FirstDayOfWeek.Monday) - 8)
                End If
                '
                ' jetzt ist tmpstart auf Montag ... 
                Dim tmpDay As Date
                Dim i As Integer, w As Integer
                i = 1
                For w = 1 To 30
                    For d = 0 To 4
                        ' das sind Montag bis Freitag
                        tmpDay = tmpstart.AddDays(d)
                        If Not feierTage.Contains(tmpDay) Then
                            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                            i = i + 1
                        End If
                    Next
                    tmpDay = tmpstart.AddDays(5)
                    If workOnSat Then
                        CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                        i = i + 1
                    End If
                    tmpDay = tmpstart.AddDays(6)
                    If workOnSun Then
                        CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                        i = i + 1
                    End If
                    tmpstart = tmpstart.AddDays(7)
                Next


            End If


            ' hier werden jetzt die Spaltenbreiten und Zeilenhöhen gesetzt 

            Dim maxRows As Integer = .Rows.Count
            Dim maxColumns As Integer = .Columns.Count


            CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
            CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe2 * 0.5

            CType(.Columns(1), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 20.0
            CType(.Columns(2), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 20.0
            CType(.Range(.Cells(1, 3), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = awinSettings.spaltenbreite


            '.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            '.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


        End With


        With Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
            .DisplayWorkbookTabs = False
            .GridlineColor = RGB(220, 220, 220)
            .FreezePanes = True
            .DisplayHeadings = True
        End With


        Application.EnableEvents = formerEE
        Application.ScreenUpdating = True

    End Sub

    Private Sub Tabelle2_Startup() Handles Me.Startup

    End Sub

    Private Sub Tabelle2_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Tabelle2_Change(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.Change

        'was_changed = True

    End Sub

    Private Sub Tabelle2_Deactivate() Handles Me.Deactivate


    End Sub

    Private Sub Tabelle2_SelectionChange(Target As Microsoft.Office.Interop.Excel.Range) Handles Me.SelectionChange

    End Sub
End Class
