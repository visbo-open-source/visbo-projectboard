Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess



Module Module1

    Sub Main(ByVal args() As String)
        For i = 0 To args.Length - 1
            Console.WriteLine("{0} => {1}", i, args(i))
        Next

        'Dim inputfile As String = My.Computer.FileSystem.CurrentDirectory & "\" & args(0)
        Dim inputfile As String = args(0)
        Dim username As String = args(1)
        Dim password As String = args(2)

        Dim xlsBatchFile As Excel.Workbook = Nothing
        Dim currentBatchfile As String

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1

        Dim reportname As String = ""
        Dim profilname As String = ""
        Dim portfolio_projname As String = ""
        Dim variantname As String = ""
        Dim rangeleft As Date
        Dim rangeright As Date

        Call MsgBox("inputfile= " & inputfile)

        Dim path As String = My.Computer.FileSystem.CurrentDirectory.ToString

        appInstance = New Microsoft.Office.Interop.Excel.Application
        Try
            If Not readawinSettings(path) Then

                awinSettings.databaseURL = My.Settings.mongoDBURL
                awinSettings.databaseName = My.Settings.mongoDBName
                awinSettings.globalPath = My.Settings.globalPath
                awinSettings.awinPath = My.Settings.awinPath

            End If

            Call awinsetTypen("ReportGen")

        Catch ex As Exception

            Call MsgBox(ex.Message)

        Finally

        End Try

        ' einlesen der Batch-Vorgabe
        xlsBatchFile = appInstance.Workbooks.Open(awinPath & inputfile, [ReadOnly]:=True, Editable:=False)
        currentBatchfile = appInstance.ActiveWorkbook.Name

        Dim wsName As Excel.Worksheet = CType(appInstance.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

        Try
            With wsName

                Dim lastrow As Integer = CType(.Cells(2000, 1), Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                Dim lastcolumn As Integer = CType(.Cells(1, 2000), Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column

                While zeile <= lastrow
                    Try

                        reportname = CStr(CType(.Cells(zeile, spalte), Microsoft.Office.Interop.Excel.Range).Value)
                        profilname = CStr(CType(.Cells(zeile, spalte + 1), Microsoft.Office.Interop.Excel.Range).Value)
                        portfolio_projname = CStr(CType(.Cells(zeile, spalte + 2), Microsoft.Office.Interop.Excel.Range).Value)
                        variantname = CStr(CType(.Cells(zeile, spalte + 3), Microsoft.Office.Interop.Excel.Range).Value)
                        If IsNothing(variantname) Then
                            variantname = ""
                        End If
                        rangeleft = CType(.Cells(zeile, spalte + 4), Microsoft.Office.Interop.Excel.Range).Value
                        rangeright = CType(.Cells(zeile, spalte + 5), Microsoft.Office.Interop.Excel.Range).Value

                        showRangeLeft = getColumnOfDate(rangeleft)
                        showRangeRight = getColumnOfDate(rangeright)

                        If Not (IsNothing(reportname) _
                            And IsNothing(profilname) _
                            And IsNothing(portfolio_projname) _
                            And IsNothing(rangeleft) _
                            And IsNothing(rangeright)) Then

                            Dim erfolgreich As Boolean = reportErstellen(portfolio_projname, variantname, profilname, reportname, username, password)
                            If erfolgreich Then
                                ' Powerpoint-Report unter dem namen reportname speichern

                            Else
                                Call logfileSchreiben("Fehler in den Angaben:  Zeile " & zeile, "Main", 0)
                            End If

                        End If


                    Catch ex As Exception
                        ' da alles im Batch ablaufen soll, soll nicht abgebrochen werden, sondern das komplette Batch-Vorgabe-File abgearbeitet
                        ' Fehlt eine Angabe, so wird ein Logbuch-Eintrag vorgenommen und die nächste Zeile des Batch-Vorgabe-Files wird ausgelesen und verarbeitet

                    End Try

                    zeile = zeile + 1

                End While

            End With
        Catch ex As Exception


        End Try

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, username, password)

    End Sub


End Module
