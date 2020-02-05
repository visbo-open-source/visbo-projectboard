

Imports ProjectBoardDefinitions
'Imports DBAccLayer
Imports ClassLibrary1
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
'Imports System.ComponentModel
'Imports System.Windows
Imports System.Windows.Forms
Imports System.Security.Principal
Imports System.Text.RegularExpressions

Public Module agm3

    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der InternenAnwesenheitslisten. 
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="kapaFile"></param>
    ''' <param name="kapaConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkCapaImportConfig(ByVal configFile As String,
                                      ByRef kapaFile As String,
                                      ByRef kapaConfigs As SortedList(Of String, clsConfigKapaImport),
                                      ByRef lastrow As Integer, ByRef oPCollection As Collection) As Boolean

        Dim outputline As String = ""
        Dim configLine As New clsConfigKapaImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim anzOld_oPCollection As Integer = oPCollection.Count
        'Dim found As Boolean
        'Dim i As Integer

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try
                configWB = appInstance.Workbooks.Open(configFile)

                Try

                    If appInstance.Worksheets.Count > 0 Then

                        currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)

                        If ok Then
                            With currentWS
                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigKapaImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "Kapa-Datei"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.capacityFile = CStr(.Cells(i, InputFileCol).value)
                                            kapaFile = configLine.capacityFile

                                        Case "month"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "year"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "role"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)


                                        Case "valueStart"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "valueLength"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "valueSign"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case "LastLine"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.column = CInt(.Cells(i, SNCol).value)
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)
                                            configLine.row = CInt(.Cells(i, ZNCol).value)
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.regex = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)

                                    End Select

                                    If kapaConfigs.ContainsKey(configLine.Titel) Then
                                        kapaConfigs.Remove(configLine.Titel)
                                    End If

                                    kapaConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With
                        Else
                            outputline = "Die Konfigurationsdatei stimmt nicht mit der erwarteten Struktur überein!"
                            oPCollection.Add(outputline)
                        End If

                    End If

                Catch ex As Exception
                    outputline = "Fehler beim Lesen der Konfigurationsdatei ..."
                    oPCollection.Add(outputline)
                End Try

                ' configCapaImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                outputline = "Die Konfigurationsdatei konnte nicht geöffnet werden - " & configFile
                oPCollection.Add(outputline)
                'Call MsgBox(outputline)
            End Try
        Else
            ' soll nur Info im Logbuch sein
            outputline = "Keine Konfigurationsdatei für Import Capacities vorhanden! - " & configFile
            Call logfileSchreiben(outputline, "", -1)
        End If

        checkCapaImportConfig = (kapaConfigs.Count > 0) And (anzOld_oPCollection - oPCollection.Count = 0)

    End Function


    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der Projekte. 
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="ProjectsFile"></param>
    ''' <param name="ProjectsConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkProjectImportConfig(ByVal configFile As String,
                                      ByRef ProjectsFile As String,
                                      ByRef ProjectsConfigs As SortedList(Of String, clsConfigProjectsImport),
                                      ByRef lastrow As Integer,
                                      ByRef outputCollection As Collection) As Boolean

        Dim configLine As New clsConfigProjectsImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim outputLine As String

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try
                configWB = appInstance.Workbooks.Open(configFile)

                Try

                    If appInstance.Worksheets.Count > 0 Then

                        currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)

                        If ok Then
                            With currentWS
                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigProjectsImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "DateiName"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.ProjectsFile = CStr(.Cells(i, InputFileCol).value)
                                            ProjectsFile = configLine.ProjectsFile

                                            'Case "ProjectTemplate"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "Budget"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "ProjectDescription"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "ProjectStart"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "ProjectEnd"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "duration"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "BU"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "ProjectNumber"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "ProjectName"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "TimeUnit"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "Ressourcen"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "days"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "weeks"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "months"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "years"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "Total"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                            'Case "LastLine"
                                            '    configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            '    configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            '    configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            '    configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            '    configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            '    configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            '    configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.column.von = CInt(hstr(0))
                                            '            configLine.column.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, SNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.column.von = CInt(.Cells(i, SNCol).value)
                                            '        configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            '    End If
                                            '    configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            '    If configLine.cellrange Then
                                            '        Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                            '        Dim hstr() As String = Split(colrange, ":")
                                            '        If hstr.Length = 2 Then
                                            '            configLine.row.von = CInt(hstr(0))
                                            '            configLine.row.bis = CInt(hstr(1))
                                            '        ElseIf hstr.Length = 1 Then
                                            '            configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '            configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '        Else
                                            '            outputLine = configLine.Titel & " : Angabe ist kein Range"
                                            '        End If
                                            '    Else
                                            '        configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                            '        configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            '    End If
                                            '    configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            '    configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            '    configLine.content = CStr(.Cells(i, InhaltCol).value)

                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)
                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                                Dim hstr() As String = Split(colrange, ":")
                                                If hstr.Length = 2 Then
                                                    configLine.column.von = CInt(hstr(0))
                                                    configLine.column.bis = CInt(hstr(1))
                                                ElseIf hstr.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, SNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                End If
                                            Else
                                                configLine.column.von = CInt(.Cells(i, SNCol).value)
                                                configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            End If
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                                Dim hstr() As String = Split(colrange, ":")
                                                If hstr.Length = 2 Then
                                                    configLine.row.von = CInt(hstr(0))
                                                    configLine.row.bis = CInt(hstr(1))
                                                ElseIf hstr.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                End If
                                            Else
                                                configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            End If
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)
                                    End Select

                                    If ProjectsConfigs.ContainsKey(configLine.Titel) Then
                                        ProjectsConfigs.Remove(configLine.Titel)
                                    End If

                                    ProjectsConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With
                        Else
                            If awinSettings.englishLanguage Then
                                outputLine = "The structure of the configFile doesn't match!  -  " & configFile
                            Else
                                outputLine = "Der Aufbau der Konfigurationsdatei ist nicht passend  -  " & configFile
                            End If
                            outputCollection.Add(outputLine)
                        End If

                    End If

                Catch ex As Exception

                End Try

                ' configCapaImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                If awinSettings.englishLanguage Then
                    Call MsgBox("The configration-file " & configFile & "  to import the projects couldn't be opened.")
                    outputLine = "The configrationfile " & configFile & "  to import the projects couldn't be opened."
                Else
                    Call MsgBox("Das Öffnen der Konfigurationsdatei " & configFile & " war nicht erfolgreich." &
                                vbCrLf & " Die Projekte können somit nicht importiert werden")
                    outputLine = "Das Öffnen der Konfigurationsdatei " & configFile & " war nicht erfolgreich." &
                                vbCrLf & " Die Projekte können somit nicht importiert werden"
                End If
                outputCollection.Add(outputLine)
            End Try
        Else
            If awinSettings.englishLanguage Then
                outputLine = "The configration-file don't exists!  -  " & configFile
            Else
                outputLine = "Die Konfigurationsdatei existiert nicht!  -  " & configFile
            End If
            outputCollection.Add(outputLine)
        End If

        checkProjectImportConfig = (ProjectsConfigs.Count > 0)

    End Function

    ''' <summary>
    ''' überprüft, ob die Voraussetzungen für das Einlesen der Projekte. 
    ''' </summary>
    ''' <param name="configFile"></param>
    ''' <param name="ActualDataFile"></param>
    ''' <param name="ActualDataConfigs"></param>
    ''' <param name="lastrow"></param>
    ''' <returns></returns>
    Public Function checkActualDataImportConfig(ByVal configFile As String,
                                      ByRef ActualDataFile As String,
                                      ByRef ActualDataConfigs As SortedList(Of String, clsConfigActualDataImport),
                                      ByRef lastrow As Integer,
                                      ByRef outputCollection As Collection) As Boolean

        Dim configLine As New clsConfigActualDataImport
        Dim currentDirectoryName As String = requirementsOrdner
        Dim configWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim outputLine As String

        ''
        '' Config-file wird geöffnet
        ' Filename ggf. mit Directory erweitern
        configFile = My.Computer.FileSystem.CombinePath(currentDirectoryName, configFile)

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(configFile) Then

            Try
                configWB = appInstance.Workbooks.Open(configFile)

                Try
                    If appInstance.Worksheets.Count > 0 Then

                        currentWS = CType(appInstance.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Dim titleCol As Integer,
                            IdentCol As Integer,
                            InputFileCol As Integer,
                            TypCol As Integer,
                            DatenCol As Integer,
                            TabUCol As Integer, TabNCol As Integer,
                            SUCol As Integer, SNCol As Integer,
                            ZUCol As Integer, ZNCol As Integer,
                            ObjCol As Integer,
                            InhaltCol As Integer

                        searcharea = currentWS.Rows(5)          ' Zeile 5 enthält die verschieden Configurationselemente

                        titleCol = searcharea.Find("Titel").Column
                        IdentCol = searcharea.Find("Identifier").Column
                        InputFileCol = searcharea.Find("InputFile").Column
                        TypCol = searcharea.Find("Typ").Column
                        DatenCol = searcharea.Find("Datenbereich").Column
                        TabUCol = searcharea.Find("Tabellen-Name").Column
                        TabNCol = searcharea.Find("Tabellen-Nummer").Column
                        SUCol = searcharea.Find("Spaltenüberschrift").Column
                        SNCol = searcharea.Find("Spalten-Nummer").Column
                        ZUCol = searcharea.Find("Zeilenbeschriftung").Column
                        ZNCol = searcharea.Find("Zeilen-Nummer").Column
                        ObjCol = searcharea.Find("Objekt-Typ").Column
                        InhaltCol = searcharea.Find("Inhalt").Column

                        Dim ok As Boolean = (titleCol + IdentCol + TypCol + DatenCol + SUCol + SNCol + ZUCol + ZNCol + ObjCol + InhaltCol > 13)

                        If ok Then
                            With currentWS
                                lastrow = .Cells(.Rows.Count, titleCol).end(Microsoft.Office.Interop.Excel.XlDirection.xlUp).row

                                For i = 6 To lastrow

                                    configLine = New clsConfigActualDataImport

                                    Dim Titel As String = CStr(.Cells(i, titleCol).value)

                                    Select Case Titel
                                        Case "DateiName"
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            ActualDataFile = configLine.Inputfile



                                        Case Else
                                            configLine.Titel = CStr(.Cells(i, titleCol).value)
                                            configLine.Identifier = CStr(.Cells(i, IdentCol).value)
                                            configLine.Inputfile = CStr(.Cells(i, InputFileCol).value)
                                            configLine.Typ = CStr(.Cells(i, TypCol).value)
                                            configLine.cellrange = (CStr(.Cells(i, DatenCol).value) = "Range")
                                            'configLine.sheet = CInt(.Cells(i, TabNCol).value)
                                            Dim tabrange As String = CStr(.Cells(i, TabNCol).value)
                                            Dim hstr() As String = Split(tabrange, ":")
                                            If hstr.Length = 2 Then
                                                configLine.sheet.von = CInt(hstr(0))
                                                configLine.sheet.bis = CInt(hstr(1))
                                            ElseIf hstr.Length = 1 Then
                                                configLine.sheet.von = CInt(.Cells(i, SNCol).value)
                                                configLine.sheet.bis = CInt(.Cells(i, SNCol).value)
                                            Else
                                                outputLine = configLine.Titel & " : Angabe für Sheet ist kein Range"
                                            End If
                                            configLine.sheetDescript = CStr(.Cells(i, TabUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, SNCol).value)
                                                Dim hstr1() As String = Split(colrange, ":")
                                                If hstr1.Length = 2 Then
                                                    configLine.column.von = CInt(hstr1(0))
                                                    configLine.column.bis = CInt(hstr1(1))
                                                ElseIf hstr1.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, SNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, SNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                End If
                                            Else
                                                configLine.column.von = CInt(.Cells(i, SNCol).value)
                                                configLine.column.bis = CInt(.Cells(i, SNCol).value)
                                            End If
                                            configLine.columnDescript = CStr(.Cells(i, SUCol).value)

                                            If configLine.cellrange Then
                                                Dim colrange As String = CStr(.Cells(i, ZNCol).value)
                                                Dim hstr2() As String = Split(colrange, ":")
                                                If hstr2.Length = 2 Then
                                                    configLine.row.von = CInt(hstr2(0))
                                                    configLine.row.bis = CInt(hstr2(1))
                                                ElseIf hstr2.Length = 1 Then
                                                    configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                    configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                                Else
                                                    outputLine = configLine.Titel & " : Angabe ist kein Range"
                                                End If
                                            Else
                                                configLine.row.von = CInt(.Cells(i, ZNCol).value)
                                                configLine.row.bis = CInt(.Cells(i, ZNCol).value)
                                            End If
                                            configLine.rowDescript = CStr(.Cells(i, ZUCol).value)
                                            configLine.objType = CStr(.Cells(i, ObjCol).value)
                                            configLine.content = CStr(.Cells(i, InhaltCol).value)
                                    End Select

                                    If ActualDataConfigs.ContainsKey(configLine.Titel) Then
                                        ActualDataConfigs.Remove(configLine.Titel)
                                    End If

                                    ActualDataConfigs.Add(configLine.Titel, configLine)

                                Next

                            End With

                        End If

                    End If

                Catch ex As Exception

                End Try

                ' configActualDataImport - Konfigurationsfile schließen
                configWB.Close(SaveChanges:=False)

            Catch ex As Exception
                Call MsgBox("Das Öffnen der " & configFile & " war nicht erfolgreich")
            End Try

        End If

        checkActualDataImportConfig = (ActualDataConfigs.Count > 0)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ActualDataConfig"></param>
    ''' <param name="tmpDatei"></param>
    ''' <param name="oPCollection"></param>
    ''' <returns></returns>
    Public Function readActualDataWithConfig(ByVal ActualDataConfig As SortedList(Of String, clsConfigActualDataImport),
                                             ByVal tmpDatei As String,
                                             ByVal IstDatenDate As Date,
                                             ByRef cacheProjekte As clsProjekteAlle,
                                             ByRef validProjectNames As SortedList(Of String, SortedList(Of String, Double())),
                                             ByRef projectRoleNames(,) As String,
                                             ByRef projectRoleValues(,,) As Double,
                                             ByRef updatedProjects As Integer,
                                             ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim outputline As String = ""
        Dim ok As Boolean = False
        Dim result As Boolean = True
        Dim actDataWB As Microsoft.Office.Interop.Excel.Workbook
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim regexpression As Regex
        Dim firstUrlTabelle As Integer
        Dim firstUrlspalte As Integer
        Dim firstUrlzeile As Integer
        Dim lastSpalte As Integer
        Dim lastZeile As Integer
        Dim hproj As New clsProjekt
        Dim pName As String = ""
        Dim anz_Proj As Integer = 0
        Dim searcharea As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim t As Integer = 0  ' tabellenIndex
        Dim hrole As clsRollenDefinition = Nothing
        Dim curmonth As Integer
        Dim lastValidMonth As Integer = getColumnOfDate(IstDatenDate)
        Dim stundenTotal As Integer = 0                     ' Stundenangabe in einer Zeile

        ' ======================
        ' vorarbeit der Definitionen geleistet
        ' ======================
        Try
            If My.Computer.FileSystem.FileExists(tmpDatei) Then
                Try
                    actDataWB = appInstance.Workbooks.Open(tmpDatei)

                    Dim vstart As clsConfigActualDataImport = ActualDataConfig("valueStart")
                    ' Auslesen erste Time-Sheet
                    firstUrlTabelle = vstart.sheet.von
                    firstUrlspalte = vstart.column.von
                    firstUrlzeile = vstart.row.von

                    ' Schleife über alle Tabellenblätter eines ausgewählten Excel-Files (hier = einer Rolle)
                    For t = 0 To vstart.sheet.bis

                        If Not IsNothing(vstart.sheet.von + t) Then
                            currentWS = CType(appInstance.Worksheets(vstart.sheet.von + t), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            If Not IsNothing(vstart.sheetDescript) Then
                                ok = (vstart.sheetDescript.Contains(currentWS.Name))
                            Else
                                ok = True
                            End If
                        End If

                        If Not ok Then
                            If awinSettings.englishLanguage Then
                                outputline = "the sheet " & currentWS.Name & " doesn't match with the configuration"
                            Else
                                outputline = "das Tabellenblatt " & currentWS.Name & " passt nicht zur Konfiguration"
                            End If
                            oPCollection.Add(outputline)
                            Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                            result = False
                            Exit For ' keine weiteren Tabellenblätter mehr lesen - Fehler aufgetreten
                        End If

                        If IsNothing(currentWS) Then
                            If awinSettings.englishLanguage Then
                                outputline = "the sheet " & vstart.sheetDescript & " doesn't exists in this workbook"
                            Else
                                outputline = "das Tabellenblatt " & vstart.sheetDescript & " ist nicht vorhanden"
                            End If
                            oPCollection.Add(outputline)
                            Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                            result = False
                        Else
                            ' passendes Worksheet gefunden

                            Try
                                ' Find Month
                                Dim monat As String = currentWS.Cells(ActualDataConfig("months").row.von, ActualDataConfig("months").column.von).value
                                Dim vglMonat As String = currentWS.Name
                                Dim validm As Boolean = (vglMonat.Contains(monat) Or monat.Contains(vglMonat))
                                ' find Year
                                Dim jahr As String = currentWS.Cells(ActualDataConfig("years").row.von, ActualDataConfig("years").column.von).value
                                Dim vglJahr As String = currentWS.Name
                                Dim validj As Boolean = (vglJahr.Contains(jahr) Or jahr.Contains(vglJahr))
                                Dim xxx As Date = "01." & monat & " " & jahr
                                curmonth = getColumnOfDate(xxx)

                            Catch ex As Exception
                                outputline = "Error looking for month/year"
                                oPCollection.Add(outputline)
                                Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                result = False
                            End Try

                            If curmonth <= lastValidMonth Then

                                ' Find Wertespalte - auf jedem Tabellenblatt evt. anders
                                Dim hspalte As String = ActualDataConfig("Total").columnDescript
                                Dim stdSpalteTotal As Integer = 0
                                Try
                                    Dim überschriftenzeile As Integer = ActualDataConfig("Überschriften").row.von
                                    searcharea = currentWS.Rows(überschriftenzeile)          ' Zeile über... enthält die verschieden Spaltendescript
                                    stdSpalteTotal = searcharea.Find(hspalte).Column
                                Catch ex As Exception
                                    If awinSettings.englishLanguage Then
                                        outputline = "in the sheet " & vstart.sheetDescript & " the value-column " & hspalte & " not found"
                                    Else
                                        outputline = "im Tabellenblatt " & vstart.sheetDescript & " konnte die WerteSpalte " & hspalte & " nicht gefunden werden"
                                    End If
                                    oPCollection.Add(outputline)
                                    Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                    result = False
                                End Try

                                ' find PersoNr
                                Dim vPersoNr As clsConfigActualDataImport = ActualDataConfig("PersonalNumber")
                                Try
                                    Dim personalNumber As String = currentWS.Cells(vPersoNr.row.von, vPersoNr.column.von).value
                                    ' find PersonalName
                                    Dim vPersoName As clsConfigActualDataImport = ActualDataConfig("PersonalName")
                                    Dim personalName As String = currentWS.Cells(vPersoName.row.von, vPersoName.column.von).value
                                    hrole = RoleDefinitions.getRoledefByEmployeeNr(personalNumber)

                                Catch ex As Exception
                                    If awinSettings.englishLanguage Then
                                        outputline = "in the sheet " & vstart.sheetDescript & "- there is something wrong with 'personal-No' or 'personal name'"
                                    Else
                                        outputline = "im Tabellenblatt " & vstart.sheetDescript & "- es gibt ein Fehler beim lesen der Personalnummer oder des Namens"
                                    End If
                                    oPCollection.Add(outputline)
                                    Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                    result = False
                                End Try

                                lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                                lastZeile = CType(currentWS.Cells(2000, firstUrlspalte), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                                If Not IsNothing(ActualDataConfig("valueEnd").rowDescript) Then
                                    Dim hzeile As String = ActualDataConfig("valueEnd").rowDescript
                                    Dim valueEndspalte As Integer = ActualDataConfig("valueEnd").column.von
                                    searcharea = currentWS.Columns(valueEndspalte)          ' in einer Spalte nach bestimmten Inhalt suchen
                                    lastZeile = searcharea.Find(hzeile).Row                 ' ZeilenNummer diesen Inhaltes merken
                                End If



                                ' alle Zeilen eines Tabellenblattes lesen
                                For z = firstUrlzeile To lastZeile

                                    stundenTotal = 0                ' zurücksetzen

                                    ' find ProjectNumber and the relevant Project
                                    Dim projektKDNr As String = ""
                                    Dim projKDNrConfig As clsConfigActualDataImport = ActualDataConfig("ProjectNumber")
                                    projektKDNr = CStr(currentWS.Cells(z, projKDNrConfig.column.von).value)

                                    If Not IsNothing(projektKDNr) Then

                                        If projKDNrConfig.objType = "RegEx" Then
                                            If Not IsNothing(projKDNrConfig.content) Then
                                                regexpression = New Regex(projKDNrConfig.content)
                                                Dim match As Match = regexpression.Match(projektKDNr)
                                                If match.Success Then
                                                    projektKDNr = match.Value
                                                Else
                                                    projektKDNr = Nothing
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "Attention: " & hrole.name & " Sheet: " & currentWS.Name & " Line: " & z.ToString & " no projectNo. given!"
                                                    Else
                                                        outputline = "Achtung: " & hrole.name & " Tabelle: " & currentWS.Name & " Zeile: " & z.ToString & " keine ProjektNr. angegeben!"
                                                    End If
                                                    oPCollection.Add(outputline)
                                                    Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                                End If
                                            End If
                                        End If
                                    End If

                                    If Not IsNothing(projektKDNr) Then

                                        Dim projektName As String = ""
                                        projektName = CStr(currentWS.Cells(z, ActualDataConfig("ProjectName").column.von).value)

                                        stundenTotal = CInt(currentWS.Cells(z, stdSpalteTotal).value)

                                        ' Check mit der Summenbildung in der Zeile
                                        Dim stdRange As Excel.Range = CType(currentWS.Range(currentWS.Cells(z, vstart.column.von + 2), currentWS.Cells(z, stdSpalteTotal - 2)), Microsoft.Office.Interop.Excel.Range)
                                        Dim stundenSumme As Integer = appInstance.WorksheetFunction.Sum(stdRange)
                                        If stundenTotal <> stundenSumme Then
                                            If awinSettings.englishLanguage Then
                                                outputline = "Attention: " & hrole.name & ": sum of the single values isn't the same as the value in '" & currentWS.Name & "' : " & hspalte
                                            Else
                                                outputline = "Achtung: " & hrole.name & "Die Summe der einzelnen Werte ist nicht gleich dem Eintrag in '" & currentWS.Name & "' : " & hspalte
                                            End If
                                            oPCollection.Add(outputline)
                                            Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                        End If

                                        Dim pvkey As String
                                        If Not IsNothing(projektName) Then
                                            pvkey = calcProjektKey(projektName, "")
                                        Else
                                            pvkey = ""
                                        End If

                                        If cacheProjekte.containsPNr(projektKDNr) Then
                                            hproj = cacheProjekte.getProjectByKDNr(projektKDNr)
                                            pName = hproj.name
                                        Else
                                            hproj = Nothing         ' Vorbesetzung

                                            Dim pNames As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(projektKDNr, err)
                                            If pNames.Count = 1 Then
                                                pName = pNames(1)

                                                Dim pname_ok As Boolean = pName = projektName
                                                ' Meldung noch ins Logbuch, wenn die Namen nicht übereinstimmen
                                                If Not pname_ok Then
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "projectname of projectNr. " & projektKDNr & " in the sheet is " & projektName & " in the DB it is " & pName
                                                    Else
                                                        outputline = "Projektname des Projektes Nr. " & projektKDNr & " in der ExcelTabelle ist " & projektName & " in der DB heißt das Projekt " & pName
                                                    End If
                                                    Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                                End If

                                                hproj = New clsProjekt
                                                hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, "", "", Date.Now, err)

                                            ElseIf pNames.Count > 1 Then
                                                ' Fehlermeldung, falls mehrer Projekte zu einer ProjektKdNr. existieren
                                                outputline = "There exists more than one project zu project No. '" & projektKDNr & "'"
                                                oPCollection.Add(outputline)
                                                Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)

                                            Else
                                                ' Fehlermeldung, falls kein Projekt zu einer ProjektKdNr. existieren
                                                outputline = "No project zu project No. '" & projektKDNr & "' User: '" & hrole.name & "' Monat: '" & currentWS.Name & "'"
                                                oPCollection.Add(outputline)
                                                Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)

                                            End If
                                        End If

                                        If IsNothing(hproj) Then
                                            'Fehler, Projekt mit einer ProjektNr. existiert in DB nicht, Keine Istdaten hierzu einlesbar
                                            'If awinSettings.englishLanguage Then
                                            '    outputline = "project Nr. " & projektKDNr & " doesn't exist in the DB. No actual data can be stored"
                                            'Else
                                            '    outputline = "Projekt mit der  Projekt-Nummer " & projektKDNr & "existiert in der DB nicht. Istdaten sind nicht zuordenbar"
                                            'End If
                                            'oPCollection.Add(outputline)
                                            'Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                            'result = False

                                        Else
                                            cacheProjekte.Add(hproj, updateCurrentConstellation:=False)                    ' Projekt in cacheProjekte merken

                                            Dim projBeginn = getColumnOfDate(hproj.startDate)
                                            Dim projEnde As Integer = getColumnOfDate(hproj.endeDate)

                                            ' Aufbauen des Eintrags
                                            Dim roleValues As New SortedList(Of String, Double())
                                            Dim tmpValues() As Double

                                            ReDim tmpValues(lastValidMonth - projBeginn)
                                            Dim teamID As Integer = -1

                                            If Not IsNothing(hrole) Then

                                                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(hrole.name, "")

                                                If Not validProjectNames.ContainsKey(pName) Then

                                                    roleValues = New SortedList(Of String, Double())

                                                    ' es handelt sich um Stunden, also in PT umrechnen 
                                                    tmpValues(curmonth - projBeginn) = stundenTotal / 8

                                                    roleValues.Add(roleNameID, tmpValues)
                                                    validProjectNames.Add(pName, roleValues)

                                                Else
                                                    roleValues = validProjectNames.Item(pName)
                                                    If roleValues.ContainsKey(roleNameID) Then
                                                        ' rolle ist bereits enthalten 
                                                        ' also summieren 
                                                        tmpValues = roleValues.Item(roleNameID)

                                                        tmpValues(curmonth - projBeginn) = tmpValues(curmonth - projBeginn) + stundenTotal / 8


                                                    Else
                                                        ' Rolle ist noch nicht enthalten 

                                                        ' es handelt sich Stunden, also in PT umrechnen 
                                                        tmpValues(curmonth - projBeginn) = stundenTotal / 8

                                                        roleValues.Add(roleNameID, tmpValues)
                                                    End If

                                                End If

                                            Else
                                                'Fehler, darf nur ein Name zu einer ProjektNr. existieren => TimeSheets nicht ins archiv
                                                outputline = "Die Rolle '" & hrole.name & "' ist nicht angelegt in Ihrem VISBO project warehouse"
                                                oPCollection.Add(outputline)
                                                result = False
                                            End If
                                        End If
                                    Else
                                        'Fehler, es ist keine ProjektKDNr angegeben, Keine Istdaten hierzu einlesbar
                                        If stundenTotal <> 0 Then
                                            If awinSettings.englishLanguage Then
                                                outputline = "there exists no project Nr. in line " & z.ToString
                                            Else
                                                outputline = "es ist keine Projekt-Nummer in Zeile " & z.ToString
                                            End If
                                            oPCollection.Add(outputline)
                                            Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                        End If
                                    End If      ' if ProjektKDNr = ""

                                Next z          'nächste Zeile lesen
                            Else
                                ' Infomeldung im Logbuch
                                outputline = "Ende der Istdaten für '" & hrole.name & "' erreicht"
                                Call logfileSchreiben(outputline, "readActualDataWithConfig", anzFehler)
                                Exit For
                            End If

                        End If

                    Next t    ' nächste Tabelle des Excel-Inputfiles

                Catch ex As Exception
                    actDataWB = Nothing
                    Call MsgBox("1 " & ex.Message)
                End Try

                If Not IsNothing(actDataWB) Then
                    actDataWB.Close(SaveChanges:=False)
                End If


            End If
        Catch ex As Exception
            Call MsgBox("2 " & ex.Message)
        End Try


        readActualDataWithConfig = result
    End Function

    Public Function readActualData(ByVal dateiname As String) As Boolean

        'dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)

        Dim oCollection As New Collection

        Try
            ' hier wird jetzt der Import gemacht 
            Call logfileSchreiben("Beginn Import Istdaten", dateiname, -1)

            ' Öffnen des Organisations-Files
            appInstance.Workbooks.Open(dateiname)
            Dim scenarioNameP As String = appInstance.ActiveWorkbook.Name



            ' das Formular aufschalten mit 
            '
            Dim editActualDataMonth As New frmProvideActualDataMonth

            If editActualDataMonth.ShowDialog = DialogResult.OK Then

                Dim monat As Integer = CInt(editActualDataMonth.valueMonth.Text)

                Dim readPastAndFutureData As Boolean = editActualDataMonth.readPastAndFutureData.Checked
                Dim createUnknownProjects As Boolean = editActualDataMonth.createUnknownProjects.Checked


                Call ImportAllianzIstdaten(monat, readPastAndFutureData, createUnknownProjects, oCollection)

            End If


            Dim wbName As String = My.Computer.FileSystem.GetName(dateiname)

            ' Schliessen des CustomUser Role-Files
            appInstance.Workbooks(wbName).Close(SaveChanges:=True)

            'sessionConstellationP enthält alle Projekte aus dem Import 
            Dim sessionConstellationP As clsConstellation = verarbeiteImportProjekte(scenarioNameP, noComparison:=False, considerSummaryProjects:=False)


            If sessionConstellationP.count > 0 Then

                If projectConstellations.Contains(scenarioNameP) Then
                    projectConstellations.Remove(scenarioNameP)
                End If

                projectConstellations.Add(sessionConstellationP)
                ' jetzt auf Projekt-Tafel anzeigen 
                Call loadSessionConstellation(scenarioNameP, False, True)

            Else
                Call MsgBox("keine Projekte importiert ...")
            End If

            If ImportProjekte.Count > 0 Then
                ImportProjekte.Clear(False)
            End If

        Catch ex As Exception

        End Try


        readActualData = oCollection.Count > 0
    End Function

    ''' <summary>
    ''' liest das im Diretory ../ressource manager evt. liegende File 'Urlaubsplaner*.xlsx' File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function readAvailabilityOfRole(ByVal kapaFileName As String, ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim old_oPCollectionCount As Integer = oPCollection.Count

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        Dim fehler As Boolean = False

        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim spalte As Integer = 2
        Dim firstUrlspalte As Integer = 5
        Dim noColor As Integer = -4142
        Dim whiteColor As Integer = 2
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date

        Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
        Dim anzMonthDays As Integer = 0
        Dim colDate As Integer = 0
        Dim anzDays As Integer = 0

        Dim lastZeile As Integer
        Dim lastSpalte As Integer
        Dim monthDays As New SortedList(Of Integer, Integer)

        Dim hrole As New clsRollenDefinition
        Dim rolename As String = ""

        Dim outPutCollection As New Collection

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(kapaFileName) Then

            Try
                kapaWB = appInstance.Workbooks.Open(kapaFileName)

                Try
                    For index = 1 To appInstance.Worksheets.Count

                        'If Not ok Then
                        '    Exit For
                        'End If


                        currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        Dim hstr() As String = Split(currentWS.Name, "Halbjahr", , )
                        If hstr.Length > 1 Then

                            ok = True
                            ' Auslesen der Jahreszahl, falls vorhanden
                            If Not IsNothing(CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                year = CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value
                            End If

                            monthDays.Clear()
                            anzDays = 0


                            lastSpalte = CType(currentWS.Cells(4, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                            lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                            ' letzte Zeile bestimmen, wenn dies verbunden Zellen sind
                            ' -------------------------------------
                            Dim rng As Range
                            Dim rngEnd As Range

                            rng = CType(currentWS.Cells(lastZeile, 1), Global.Microsoft.Office.Interop.Excel.Range)

                            If rng.MergeCells Then

                                rng = rng.MergeArea
                                rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

                                ' dann ist die lastZeile neu zu besetzen
                                lastZeile = rngEnd.Row
                            End If

                            ' nun hat die Variable lastZeile sicher den richtigen Wert
                            ' --------------------------------------


                            Dim vglColor As Integer = noColor         ' keine Farbe
                            Dim i As Integer = firstUrlspalte

                            While ok And i <= lastSpalte

                                If vglColor <> CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex Then
                                    ok = (anzDays = anzMonthDays) Or (anzDays = 0)
                                    vglColor = CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex
                                    anzDays = 1
                                Else
                                    If CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Text <> "" Then
                                        Dim monthName As String = CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Text
                                        ' ''Dim strDate As String = "01." & monthName & " " & year
                                        ' ''Dim hdate As DateTime = DateValue(strDate)

                                        Dim isdate As Boolean = DateTime.TryParse(monthName & " " & year.ToString, tmpDate)
                                        If isdate Then
                                            colDate = getColumnOfDate(tmpDate)
                                            anzMonthDays = DateTime.DaysInMonth(year, Month(tmpDate))
                                            monthDays.Add(colDate, anzMonthDays)
                                        End If
                                    End If

                                    anzDays = anzDays + 1
                                End If

                                i = i + 1
                            End While


                            If Not ok Then

                                fehler = True

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Error reading planning holidays: Please check the calendar in this file ..."
                                Else
                                    msgtxt = "Fehler beim Lesen der Urlaubsplanung: Bitte prüfen Sie die Korrektheit des Kalenders ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                'Call MsgBox(msgtxt)

                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)

                                If formerEE Then
                                    appInstance.EnableEvents = True
                                End If

                                If formerSU Then
                                    appInstance.ScreenUpdating = True
                                End If

                                enableOnUpdate = True
                                If awinSettings.englishLanguage Then
                                    msgtxt = "Your planning holidays couldn't be read, because of problems"
                                Else
                                    msgtxt = "Ihre Urlaubsplanung konnte nicht berücksichtigt werden"
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If

                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                'Call showOutPut(oPCollection, "Lesen Urlaubsplanung wurde mit Fehler abgeschlossen", "Meldungen zu Lesen Urlaubsplanung")
                                ' tk 12.2.19 ess oll alles gelesen werden - es wird nicht weitergemacht, wenn es Einträge in der outputCollection gibt 
                                'Throw New ArgumentException(msgtxt)
                            Else

                                For iZ = 5 To lastZeile


                                    rolename = CType(currentWS.Cells(iZ, 2), Global.Microsoft.Office.Interop.Excel.Range).Text
                                    If rolename <> "" Then
                                        hrole = RoleDefinitions.getRoledef(rolename)
                                        If Not IsNothing(hrole) Then

                                            Dim defaultHrsPerdayForThisPerson As Double = hrole.defaultDayCapa

                                            Dim iSp As Integer = firstUrlspalte
                                            Dim anzArbTage As Double = 0
                                            Dim anzArbStd As Double = 0

                                            For Each kvp As KeyValuePair(Of Integer, Integer) In monthDays

                                                Dim colOfDate As Integer = kvp.Key
                                                anzDays = kvp.Value
                                                For sp = iSp + 0 To iSp + anzDays - 1

                                                    If iSp <= lastSpalte Then
                                                        Dim hint As Integer = CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex)

                                                        If CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex) = noColor _
                                                            Or CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex = whiteColor Then

                                                            If Not IsNothing(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                If IsNumeric(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                    Dim angabeInStd As Double = CType(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value, Double)

                                                                    If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                        anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                    Else
                                                                        If awinSettings.englishLanguage Then
                                                                            msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                        Else
                                                                            msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                        End If
                                                                        If Not oPCollection.Contains(msgtxt) Then
                                                                            oPCollection.Add(msgtxt, msgtxt)
                                                                        End If
                                                                        'Call MsgBox(msgtxt)
                                                                        fehler = True
                                                                        Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                    End If
                                                                Else
                                                                    ' Feld ist weiss, oder hat keine Farbe, keine Zahl: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                End If



                                                            Else

                                                                ' hier wird die Telair Variante gemacht 
                                                                ' das einfachste wäre eigentlich  
                                                                'anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex = noColor Then
                                                                    'anzArbStd = anzArbStd + 8
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                Else
                                                                    ' freier Tag für Teilzeitbeschäftigte
                                                                    msgtxt = "Tag zählt nicht: Zeile " & iZ & ", Spalte " & sp
                                                                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                End If

                                                            End If
                                                        End If
                                                    Else
                                                        If awinSettings.englishLanguage Then
                                                            msgtxt = "Error reading the amount of working days of " & hrole.name & " ..."
                                                        Else
                                                            msgtxt = "Fehler beim Lesen der verfügbaren Arbeitstage von " & hrole.name & " ..."
                                                        End If
                                                        fehler = True
                                                        If Not oPCollection.Contains(msgtxt) Then
                                                            oPCollection.Add(msgtxt, msgtxt)
                                                        End If
                                                        Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                    End If

                                                Next

                                                anzArbTage = anzArbStd / 8

                                                'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                                                If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                    hrole.kapazitaet(colOfDate) = anzArbTage
                                                Else
                                                    hrole.kapazitaet(colOfDate) = 0
                                                End If

                                                iSp = iSp + anzDays
                                                anzArbTage = 0              ' Anzahl Arbeitstage wieder zurücksetzen für den nächsten Monat
                                                anzArbStd = 0               ' Anzahl zu leistender Arbeitsstunden wieder zurücksetzen für den nächsten Monat

                                            Next

                                        Else

                                            If awinSettings.englishLanguage Then
                                                msgtxt = "Role " & rolename & " not defined ..."
                                            Else
                                                msgtxt = "Rolle " & rolename & " nicht definiert ..."
                                            End If
                                            If Not oPCollection.Contains(msgtxt) Then
                                                oPCollection.Add(msgtxt, msgtxt)
                                            End If
                                            'Call MsgBox(msgtxt)
                                            fehler = True
                                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                        End If
                                    Else

                                        If awinSettings.englishLanguage Then
                                            msgtxt = "No Name of role given ..."
                                        Else
                                            msgtxt = "kein Rollenname angegeben ..."
                                        End If
                                        If Not oPCollection.Contains(msgtxt) Then
                                            oPCollection.Add(msgtxt, msgtxt)
                                        End If
                                        Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                    End If

                                Next iZ

                            End If   ' ende von if not OK
                        Else
                            If awinSettings.visboDebug Then

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Worksheet " & hstr(0) & "doesn't belongs to planning holidays ..."
                                Else
                                    msgtxt = "Worksheet" & hstr(0) & " gehört nicht zum Urlaubsplaner ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                            End If

                        End If

                    Next index


                Catch ex2 As Exception
                    'If fehler Then
                    '    'Call MsgBox(msgtxt)

                    '    RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(DateTime.Now, err)

                    '    msgtxt = "Es wurden nun die Kapazitäten aus der Datenbank gelesen ..."
                    '    If awinSettings.englishLanguage Then
                    '        msgtxt = "Therefore read the capacity of every Role from the DB  ..."
                    '    End If
                    '    If Not oPCollection.Contains(msgtxt) Then
                    '        oPCollection.Add(msgtxt, msgtxt)
                    '    End If
                    '    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                    'End If
                End Try

                'kapaWB.Close(SaveChanges:=False)
            Catch ex As Exception

            End Try

        End If


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True
        kapaWB.Close(SaveChanges:=False)

        ' das wird jetzt an der übergeordneten Stelle gemacht
        'Call showOutPut(oPCollection, "Meldungen zu Lesen Urlaubsplanung", "Folgende Probleme sind beim Lesen der Urlaubsplanung aufgetreten")

        ' ''If outPutCollection.Count > 0 Then
        ' ''    Call showOutPut(outPutCollection, _
        ' ''                    "Meldungen Einlesevorgang Urlaubsdatei", _
        ' ''                    "zum Zeitpunkt " & storedAtOrBefore.ToString & " aufgeführte Rolle nicht definiert")
        ' ''End If

        readAvailabilityOfRole = (oPCollection.Count = old_oPCollectionCount)

    End Function

    ''' <summary>
    ''' liest das im Diretory ../ressource manager evt. liegende File 'zeuss*.xlsx' (oder wie in kapaConfig benamst) File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function readAvailabilityOfRoleWithConfig(ByVal kapaConfig As SortedList(Of String, clsConfigKapaImport),
                                                ByVal kapaFileName As String,
                                                ByRef oPCollection As Collection) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim old_oPCollectionCount As Integer = oPCollection.Count

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        Dim fehler As Boolean = False

        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim spalte As Integer = 2
        Dim firstUrlspalte As Integer = 0
        Dim firstUrlzeile As Integer = 0
        Dim noColor As Integer = -4142
        Dim whiteColor As Integer = 2
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim dateConsidered As Date

        'Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
        Dim monthName As String = ""

        ' tk wird nicht verwendet ... 
        'Dim monthNumber As Integer = 0

        Dim Jahr As Integer = 0
        Dim anzMonthDays As Integer = 0
        Dim colDate As Integer = 0
        Dim anzDays As Integer = 0

        Dim lastZeile As Integer
        Dim lastSpalte As Integer
        Dim monthDays As New SortedList(Of Integer, Integer)

        Dim hrole As New clsRollenDefinition
        Dim rolename As String = ""

        Dim regexpression As Regex

        Dim outPutCollection As New Collection

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(kapaFileName) Then

            Try
                kapaWB = appInstance.Workbooks.Open(kapaFileName)

                Try
                    For index = 1 To appInstance.Worksheets.Count

                        currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        With currentWS

                            'Dim regex As String = kapaConfig("month").regex
                            'Dim Inhalt As String = kapaConfig("month").content

                            ' Auslesen der Jahreszahl, falls vorhanden
                            Dim hjahr As String = CStr(.Cells(kapaConfig("year").row, kapaConfig("year").column).value)
                            If IsNothing(hjahr) Then
                                Jahr = 0
                            Else
                                If kapaConfig("year").regex = "RegEx" Then
                                    'regexpression = New Regex("[0-9]{4}")
                                    regexpression = New Regex(kapaConfig("year").content)
                                    Dim match As Match = regexpression.Match(hjahr)
                                    If match.Success Then
                                        Jahr = CInt(match.Value)
                                    Else
                                        Jahr = 0
                                    End If
                                End If
                            End If


                            ' Auslesen des relevanten Monats
                            Dim hmonth As String = CStr(.Cells(kapaConfig("month").row, kapaConfig("month").column).value)
                            If IsNothing(hmonth) Then
                                monthName = ""
                            Else
                                If kapaConfig("month").regex = "RegEx" Then
                                    regexpression = New Regex(kapaConfig("month").content)
                                    Dim Match As Match = regexpression.Match(hmonth)
                                    If Match.Success Then
                                        monthName = Match.Value
                                    Else
                                        monthName = ""
                                    End If
                                End If
                            End If


                            ' Auslesen erste Verfügbarkeitsspalte
                            firstUrlspalte = kapaConfig("valueStart").column
                            firstUrlzeile = kapaConfig("valueStart").row
                        End With

                        ' tk 3.2.20 
                        Dim isdate As Boolean = DateTime.TryParse(monthName & " " & Jahr.ToString, dateConsidered)

                        Dim beginningDay As Integer = -1
                        Dim endingDay As Integer = -1
                        Try
                            ' das kann schiefgehen, wenn keine Zahl im Feld steht ... 
                            beginningDay = CInt(currentWS.Cells(firstUrlzeile - 1, firstUrlspalte).value)
                        Catch ex As Exception
                            beginningDay = -1
                        End Try

                        If beginningDay <> 1 Then
                            If awinSettings.englishLanguage Then
                                msgtxt = "Error in date definition row in Capa file: File, Row, Column: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte & " does not start with 1"
                            Else
                                msgtxt = "Fehler in Datums-Zeile in Kapazitäts-Datei: Datei, Zeile, Spalte: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte & "startet nicht bei 1"

                            End If
                            oPCollection.Add(msgtxt)
                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)

                        ElseIf Not isdate Then

                            If awinSettings.englishLanguage Then
                                msgtxt = "Error in Month of capacity definition: no valid month, year in Capa file: " & kapaFileName
                            Else
                                msgtxt = "Fehler in Angabe des auszulesenden Monats in Kapazitäts-Datei: " & kapaFileName

                            End If
                            oPCollection.Add(msgtxt)
                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                        Else
                            If Jahr <> 0 And monthName <> "" Then

                                colDate = getColumnOfDate(dateConsidered)

                                monthDays.Clear()

                                anzMonthDays = DateTime.DaysInMonth(Jahr, Month(dateConsidered))
                                If Not monthDays.ContainsKey(colDate) Then
                                    monthDays.Add(colDate, anzMonthDays)
                                End If

                                ' tk prüfen, ob der letzte Tag auch der richtige ist ... 
                                Try
                                    ' das kann schiefgehen, wenn keine Zahl im Feld steht ... 
                                    endingDay = CInt(currentWS.Cells(firstUrlzeile - 1, firstUrlspalte + anzMonthDays - 1).value)
                                Catch ex As Exception
                                    endingDay = -1
                                End Try

                                If endingDay <> anzMonthDays Then
                                    ok = False

                                    If awinSettings.englishLanguage Then
                                        msgtxt = "Error in date definition row in Capa file: File, Row, Column: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte + anzMonthDays - 1 & " does not show last day in month"
                                    Else
                                        msgtxt = "Fehler in Datums-Zeile in Kapazitäts-Datei: Datei, Zeile, Spalte: " & vbLf & kapaFileName & ", " & firstUrlzeile & ", " & firstUrlspalte + anzMonthDays - 1 & "zeigt nicht den letzten Tag des Monats "

                                    End If

                                    oPCollection.Add(msgtxt)
                                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)

                                Else
                                    ' hier ist sichergestellt, dass die erste Spalte mit 1 beginnt, die letzte Spalte dem Tag entspricht, mit dem der Monat endet

                                    ok = True

                                    anzDays = 0

                                    lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                                    lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                                    ' Nachkorrektur gemäss Angabe in KonfigDate 'LastLine'
                                    Dim found As Boolean = False
                                    Dim i As Integer = lastZeile + 1
                                    While Not found
                                        i = i - 1
                                        If kapaConfig("LastLine").regex = "RegEx" Then
                                            regexpression = New Regex(kapaConfig("LastLine").content)
                                            Dim lastLineContent As String = CStr(currentWS.Cells(i, kapaConfig("LastLine").column).value)
                                            If Not IsNothing(lastLineContent) Then
                                                Dim match As Match = regexpression.Match(lastLineContent)
                                                If match.Success Then
                                                    lastLineContent = match.Value
                                                    found = True
                                                End If
                                            End If
                                        End If

                                    End While
                                    lastZeile = i - 1


                                    ' letzte Zeile bestimmen, wenn dies verbunden Zellen sind
                                    ' -------------------------------------
                                    Dim rng As Range
                                    Dim rngEnd As Range

                                    rng = CType(currentWS.Cells(lastZeile, 1), Global.Microsoft.Office.Interop.Excel.Range)

                                    If rng.MergeCells Then

                                        rng = rng.MergeArea
                                        rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

                                        ' dann ist die lastZeile neu zu besetzen
                                        lastZeile = rngEnd.Row
                                    End If
                                End If


                                If Not ok Then

                                    fehler = True

                                    If awinSettings.englishLanguage Then
                                        msgtxt = "Error reading availabilities: Please check the calendar in this file ..."
                                    Else
                                        msgtxt = "Fehler beim Lesen der Verfügbarkeiten: Bitte prüfen Sie die Korrektheit des Kalenders ..."
                                    End If
                                    If Not oPCollection.Contains(msgtxt) Then
                                        oPCollection.Add(msgtxt, msgtxt)
                                    End If
                                    'Call MsgBox(msgtxt)

                                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)

                                    If formerEE Then
                                        appInstance.EnableEvents = True
                                    End If

                                    If formerSU Then
                                        appInstance.ScreenUpdating = True
                                    End If

                                    enableOnUpdate = True
                                    If awinSettings.englishLanguage Then
                                        msgtxt = "Your availabilities couldn't be read, because of problems"
                                    Else
                                        msgtxt = "Ihre Verfügbarkeiten konnten nicht berücksichtigt werden"
                                    End If
                                    If Not oPCollection.Contains(msgtxt) Then
                                        oPCollection.Add(msgtxt, msgtxt)
                                    End If

                                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                    'Call showOutPut(oPCollection, "Lesen Urlaubsplanung wurde mit Fehler abgeschlossen", "Meldungen zu Lesen Urlaubsplanung")
                                    ' tk 12.2.19 ess oll alles gelesen werden - es wird nicht weitergemacht, wenn es Einträge in der outputCollection gibt 
                                    'Throw New ArgumentException(msgtxt)
                                Else

                                    For iZ = firstUrlzeile To lastZeile


                                        rolename = CType(currentWS.Cells(iZ, kapaConfig("role").column), Global.Microsoft.Office.Interop.Excel.Range).Text

                                        ' tk 31.1.2020 Test - der CheckWert steht auf Spalte "AS"
                                        ' dazu muss manuell der Check-Wert bestimmt und in der Excel Datei eingetragen werden ..  
                                        Dim checkWert As Double = -1
                                        Try
                                            If Not IsNothing(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                If IsNumeric(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                    checkWert = CDbl(CType(currentWS.Cells(iZ, "AS"), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                End If
                                            End If
                                        Catch ex As Exception
                                            checkWert = -1
                                        End Try
                                        ' Ende tk 31.1.2020 Auslesen Checkwert für Kapa-Bestimmung 

                                        If rolename <> "" Then
                                            hrole = RoleDefinitions.getRoledef(rolename)
                                            If Not IsNothing(hrole) Then

                                                Dim defaultHrsPerdayForThisPerson As Double = hrole.defaultDayCapa

                                                Dim iSp As Integer = firstUrlspalte
                                                Dim anzArbTage As Double = 0
                                                Dim anzArbStd As Double = 0

                                                For Each kvp As KeyValuePair(Of Integer, Integer) In monthDays

                                                    Dim colOfDate As Integer = kvp.Key
                                                    anzDays = kvp.Value
                                                    For sp = iSp + 0 To iSp + anzDays - 1


                                                        If iSp <= lastSpalte Then

                                                            Dim hint As Integer = CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex)

                                                            If CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex) = noColor _
                                                                Or CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex = whiteColor Then

                                                                Dim aktCell As Object = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value

                                                                If Not IsNothing(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                    If IsNumeric(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                        Dim angabeInStd As Double = CType(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value, Double)

                                                                        If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                            anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                        Else
                                                                            If awinSettings.englishLanguage Then
                                                                                msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                            Else
                                                                                msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                            End If
                                                                            If Not oPCollection.Contains(msgtxt) Then
                                                                                oPCollection.Add(msgtxt, msgtxt)
                                                                            End If
                                                                            'Call MsgBox(msgtxt)
                                                                            fehler = True
                                                                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                        End If
                                                                    ElseIf (CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value = "") Then

                                                                        ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                        anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                    End If

                                                                Else
                                                                    ' ur:07.01.2020: Telair Variante entfällt mit Zeuss-Anpassung

                                                                    ' Feld ist ohne Inhalt: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                    '' hier wird die Telair Variante gemacht 
                                                                    '' das einfachste wäre eigentlich  
                                                                    ''anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                    ''Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                    '' ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                    ''If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value <> "/" Then
                                                                    ''    'anzArbStd = anzArbStd + 8
                                                                    ''    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                    ''Else
                                                                    ''    ' freier Tag für Teilzeitbeschäftigte
                                                                    ''    msgtxt = "Tag zählt nicht: Zeile " & iZ & ", Spalte " & sp
                                                                    ''    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                    ''End If

                                                                End If
                                                            End If
                                                        Else
                                                            If awinSettings.englishLanguage Then
                                                                msgtxt = "Error reading the amount of working days of " & hrole.name & " ..."
                                                            Else
                                                                msgtxt = "Fehler beim Lesen der verfügbaren Arbeitstage von " & hrole.name & " ..."
                                                            End If
                                                            fehler = True
                                                            If Not oPCollection.Contains(msgtxt) Then
                                                                oPCollection.Add(msgtxt, msgtxt)
                                                            End If
                                                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                        End If

                                                    Next

                                                    anzArbTage = anzArbStd / 8

                                                    ' tk 31.1.20 Check den Wert
                                                    Dim formerVD As Boolean = awinSettings.visboDebug
                                                    awinSettings.visboDebug = True
                                                    If awinSettings.visboDebug Then
                                                        If checkWert <> -1 Then
                                                            If Math.Abs(anzArbTage - checkWert) > 0.0001 Then
                                                                Call MsgBox("Abweichung in Kapa-Bestimmung")
                                                            End If
                                                        End If
                                                    End If
                                                    awinSettings.visboDebug = formerVD
                                                    'Ende tk Check den Wert 

                                                    'nur wenn die hrole schon eingetreten und nicht ausgetreten ist, wird die Capa eingetragen
                                                    If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                        hrole.kapazitaet(colOfDate) = anzArbTage
                                                    Else
                                                        hrole.kapazitaet(colOfDate) = 0
                                                    End If
                                                    iSp = iSp + anzDays
                                                    anzArbTage = 0              ' Anzahl Arbeitstage wieder zurücksetzen für den nächsten Monat
                                                    anzArbStd = 0               ' Anzahl zu leistender Arbeitsstunden wieder zurücksetzen für den nächsten Monat

                                                Next

                                            Else

                                                If awinSettings.englishLanguage Then
                                                    msgtxt = "Role " & rolename & " not defined ..."
                                                Else
                                                    msgtxt = "Rolle " & rolename & " nicht definiert ..."
                                                End If
                                                If Not oPCollection.Contains(msgtxt) Then
                                                    oPCollection.Add(msgtxt, msgtxt)
                                                End If
                                                'Call MsgBox(msgtxt)
                                                fehler = True
                                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                            End If
                                        Else

                                            If awinSettings.englishLanguage Then
                                                msgtxt = "No Name of role given ..."
                                            Else
                                                msgtxt = "kein Rollenname angegeben ..."
                                            End If
                                            If Not oPCollection.Contains(msgtxt) Then
                                                oPCollection.Add(msgtxt, msgtxt)
                                            End If
                                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                        End If

                                    Next iZ

                                End If   ' ende von if not OK
                            Else

                                If awinSettings.visboDebug Then

                                    If awinSettings.englishLanguage Then
                                        msgtxt = "Worksheet " & kapaFileName & "doesn't contain month/year ..."
                                    Else
                                        msgtxt = "Worksheet" & kapaFileName & " enthält keine Angaben zu Monat/Jahr ..."
                                    End If
                                    If Not oPCollection.Contains(msgtxt) Then
                                        oPCollection.Add(msgtxt, msgtxt)
                                    End If
                                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                End If

                            End If

                        End If      'beginningDay = 1

                    Next index


                Catch ex2 As Exception
                    If awinSettings.englishLanguage Then
                        msgtxt = "Error reading dates like month/year ..."
                    Else
                        msgtxt = "Fehler beim Lesen der notwendigen Randdaten wie Monat/Jahr ..."
                    End If
                    If Not oPCollection.Contains(msgtxt) Then
                        oPCollection.Add(msgtxt, msgtxt)
                    End If
                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                End Try

        'kapaWB.Close(SaveChanges:=False)
        Catch ex As Exception

            End Try

        End If


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True

        kapaWB.Close(SaveChanges:=False)

        readAvailabilityOfRoleWithConfig = (oPCollection.Count = old_oPCollectionCount)

    End Function

    ''' <summary>
    ''' liest Projekte gemäß Konfiguration ein 
    ''' </summary>
    ''' <param name="listOfProjectFiles"></param>
    ''' <param name="projectConfig"></param>
    ''' <param name="meldungen"></param>
    ''' <returns></returns>
    Public Function readProjectsAllg(ByVal listOfProjectFiles As Collection,
                                     ByVal projectConfig As SortedList(Of String, clsConfigProjectsImport),
                                     ByRef meldungen As Collection) As List(Of String)

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim listOfArchivFiles As New List(Of String)
        Dim anzFehler As Integer = 0
        Dim result As Boolean = False

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False



        If listOfProjectFiles.Count > 0 Then
            ' Öffnen des projectFile
            For Each tmpDatei As String In listOfProjectFiles
                Call logfileSchreiben("Einlesen Projekte " & tmpDatei, "", anzFehler)
                result = readProjectsWithConfig(projectConfig, tmpDatei, meldungen)

                If result Then
                    ' hier: merken der erfolgreich importierten Projects Dateien
                    listOfArchivFiles.Add(tmpDatei)
                End If
            Next

        Else
            Dim errMsg As String = "Es gibt keine Datei zur Projekt-Anlage" & vbLf _
                             & "Es wurden daher jetzt keine berücksichtigt"

            ' das sollte nicht dazu führen, dass nichts gemacht wird 
            'meldungen.Add(errMsg)
            'ur: 08.01.2020: endgültige meldung erst nachdem alle abgearbeitet wurden
            'Call MsgBox(errMsg)

            Call logfileSchreiben(errMsg, "", anzFehler)
        End If

        If result Then
            readProjectsAllg = listOfArchivFiles
        Else
            readProjectsAllg = New List(Of String)
        End If

    End Function
    Function readProjectsWithConfig(ByVal projectConfig As SortedList(Of String, clsConfigProjectsImport),
                                    ByVal tmpDatei As String,
                                    ByRef meldungen As Collection) As Boolean
        Dim outputline As String = ""
        Dim ok As Boolean = False
        Dim result As Boolean = False
        Dim projectWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim currentWS As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim regexpression As Regex
        Dim firstUrlspalte As Integer
        Dim firstUrlzeile As Integer
        Dim lastSpalte As Integer
        Dim lastZeile As Integer
        Dim anz_Proj As Integer = 0

        ' Variables to create a Project
        Dim hproj As clsProjekt
        Dim pName As String = ""
        Dim vName As String = ""
        Dim vorlagenName As String = ""
        Dim startDate As Date
        Dim endDate As Date
        Dim budget As Double
        Dim sfit As Double
        Dim risk As Double
        Dim projectNummer As String = ""
        Dim description As String = ""
        Dim listOfCustomFields As New Collection
        Dim businessUnit As String = ""
        Dim responsible As String = ""
        Dim status As String = ""
        Dim zeile As Integer = 0
        Dim roleNames() As String = Nothing
        Dim roleValues() As Double = Nothing
        Dim roleListNameValues As New SortedList(Of String, Double())
        Dim costNames() As String = Nothing
        Dim costValues() As Double = Nothing
        Dim phNames() As String
        Dim przPhasenAnteile() As Double
        Dim combinedName As Boolean = False
        Dim createBudget As Boolean = True
        Dim createCostsRolesAnyhow As Boolean = True

        Dim monthVon As Integer = 0
        Dim monthBis As Integer = 0

        Dim noGo As Integer = 0   'Sobald diese Variable > 0 ist, wird das Projekt nicht importiert


        Try
            If My.Computer.FileSystem.FileExists(tmpDatei) Then

                Try

                    projectWB = appInstance.Workbooks.Open(tmpDatei)

                    Dim vstart As clsConfigProjectsImport = projectConfig("valueStart")
                    ' Auslesen erste Projekt-Spalte
                    firstUrlspalte = vstart.column.von
                    firstUrlzeile = vstart.row.von

                    If appInstance.Worksheets.Count > 0 Then

                        If Not IsNothing(vstart.sheet) Then
                            currentWS = CType(appInstance.Worksheets(vstart.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            ok = (currentWS.Name = vstart.sheetDescript)
                        End If
                        If Not ok Then
                            If Not IsNothing(vstart.sheetDescript) Then
                                currentWS = CType(appInstance.Worksheets(vstart.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                            Else
                                currentWS = Nothing
                            End If
                        End If

                        If IsNothing(currentWS) Then
                            outputline = "The Worksheet you want to import cannot be matched"
                            meldungen.Add(outputline)
                            Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                        Else

                            lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                            lastZeile = CType(currentWS.Cells(2000, firstUrlspalte), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                            Try
                                Dim projNumber As String = ""

                                For i = firstUrlzeile To lastZeile + 1

                                    'Find ProjectNumber
                                    Dim projNumber_new As String = ""
                                    Try
                                        Dim projNrConfig As clsConfigProjectsImport = projectConfig("ProjectNumber")

                                        If currentWS.Index <> projNrConfig.sheet Then
                                            If Not IsNothing(projNrConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projNrConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projNrConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS

                                            Select Case projNrConfig.Typ
                                                Case "Text"
                                                    projNumber_new = CStr(.Cells(i, projNrConfig.column.von).value)
                                                Case "Integer"
                                                    projNumber_new = CInt(.Cells(i, projNrConfig.column.von).value)
                                                Case "Decimal"
                                                    projNumber_new = CDbl(.Cells(i, projNrConfig.column.von).value)
                                                Case "Date"
                                                    projNumber_new = CDate(.Cells(i, projNrConfig.column.von).value)
                                                Case Else
                                                    projNumber_new = .Cells(i, projNrConfig.column.von).value
                                            End Select

                                            If projNrConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projNrConfig.content)
                                                Dim match As Match = regexpression.Match(projNumber_new)
                                                If match.Success Then
                                                    projNumber_new = match.Value
                                                End If
                                            End If

                                        End With



                                        If projNumber_new <> projNumber And i > firstUrlzeile Then
                                            If noGo > 0 Then
                                                If awinSettings.englishLanguage Then
                                                    outputline = "Error : Project '" & pName & "' starting at: " & startDate.ToString & " finishing at: " & endDate.ToString & "  N O T  imported !"
                                                Else
                                                    outputline = "Fehler : Projekt '" & pName & "' mit Start: " & startDate.ToString & " und Ende: " & endDate.ToString & "  N I C H T  erzeugt !"
                                                End If
                                                meldungen.Add(outputline)
                                                Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)

                                                ' nach Projekt-Speicherung in ImportProjekte muss Bedarfsliste zurückgesetzt werden
                                                roleListNameValues = New SortedList(Of String, Double())

                                                ' zurücksetzen der Variable, die anzeigt, dass das aktuelle Projekt echte Fehler hatte beim Einlesen
                                                noGo = 0
                                            Else

                                                Dim anzRoles As Integer = roleListNameValues.Count
                                                ReDim roleNames(anzRoles - 1)
                                                ReDim roleValues(monthBis - monthVon)
                                                Dim k As Integer = 0
                                                For Each kvp As KeyValuePair(Of String, Double()) In roleListNameValues
                                                    roleNames(k) = kvp.Key
                                                    k = k + 1
                                                Next

                                                ReDim phNames(1)
                                                ReDim przPhasenAnteile(1)

                                                'erstelleProjektausParametern()
                                                anz_Proj = anz_Proj + 1
                                                hproj = New clsProjekt
                                                hproj = erstelleProjektausParametern(pName, vName, vorlagenName,
                                                                 startDate, endDate,
                                                                 budget, sfit, risk,
                                                                 projNumber, description,
                                                                 listOfCustomFields, businessUnit, responsible,
                                                                 status, zeile,
                                                                 roleNames, roleValues,
                                                                 costNames, costValues, phNames, przPhasenAnteile, combinedName, createBudget, createCostsRolesAnyhow)

                                                For Each kvp As KeyValuePair(Of String, Double()) In roleListNameValues

                                                    Dim tmpRCnameID As String = RoleDefinitions.bestimmeRoleNameID(kvp.Key, "")
                                                    hproj.AllPhases(0).getRoleByRoleNameID(tmpRCnameID).Xwerte = kvp.Value

                                                    Dim hilfe As Boolean = True
                                                Next

                                                ' Budget setzen 
                                                Call hproj.setBudgetAsNeeded()

                                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

                                                outputline = "Projekt '" & pName & "' mit Start: " & startDate.ToString & " und Ende: " & endDate.ToString & " erzeugt !"
                                                meldungen.Add(outputline)
                                                Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)

                                                ' nach Projekt-Speicherung in ImportProjekte muss Bedarfsliste zurückgesetzt werden
                                                roleListNameValues = New SortedList(Of String, Double())

                                            End If

                                        End If


                                        If i > lastZeile And IsNothing(projNumber_new) Then
                                            ' am Ende der zu lesenden Zeilen angekommen, die Felder sind nun leer
                                            ' beenden der Einlese-Aktion indem die For-Schleife abgebrochen wird
                                            Exit For
                                        End If

                                        projNumber = projNumber_new

                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectnumber in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden der ProjektNummer in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    'Find BusinesssUnit
                                    Dim projBU As Object
                                    Try
                                        Dim projBUConfig As clsConfigProjectsImport = projectConfig("BU")

                                        If currentWS.Index <> projBUConfig.sheet Then
                                            If Not IsNothing(projBUConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projBUConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projBUConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If
                                        With currentWS

                                            Select Case projBUConfig.Typ
                                                Case "Text"
                                                    projBU = CStr(.Cells(i, projBUConfig.column.von).value)
                                                Case "Integer"
                                                    projBU = CInt(.Cells(i, projBUConfig.column.von).value)
                                                Case "Decimal"
                                                    projBU = CDbl(.Cells(i, projBUConfig.column.von).value)
                                                Case "Color"
                                                    projBU = CLng(.Cells(i, projBUConfig.column.von).value)
                                                Case Else
                                                    projBU = .Cells(i, projBUConfig.column.von).value
                                            End Select

                                            If projBUConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projBUConfig.content)
                                                Dim match As Match = regexpression.Match(projBU)
                                                If match.Success Then
                                                    projBU = match.Value
                                                End If
                                            End If
                                            businessUnit = projBU
                                        End With
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the BU in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden der BU in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try


                                    'Find ProjectName
                                    Dim projName As String
                                    Try
                                        Dim projNameConfig As clsConfigProjectsImport = projectConfig("ProjectName")
                                        If currentWS.Index <> projNameConfig.sheet Then
                                            If Not IsNothing(projNameConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projNameConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projNameConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        With currentWS
                                            Select Case projNameConfig.Typ
                                                Case "Text"
                                                    projName = CStr(.Cells(i, projNameConfig.column.von).value)
                                                Case "Integer"
                                                    projName = CInt(.Cells(i, projNameConfig.column.von).value)
                                                Case "Decimal"
                                                    projName = CDbl(.Cells(i, projNameConfig.column.von).value)
                                                Case "Date"
                                                    projName = CDate(.Cells(i, projNameConfig.column.von).value)
                                                Case Else
                                                    projName = .Cells(i, projNameConfig.column.von).value
                                            End Select

                                            If projNameConfig.objType = "RegEx" Then
                                                regexpression = New Regex(projNameConfig.content)
                                                Dim match As Match = regexpression.Match(projName)
                                                If match.Success Then
                                                    projName = match.Value
                                                End If
                                            End If
                                            pName = projName
                                            ' ggfs. vorhandene Sonderzeichen wie (,),# [,] ersetzen
                                            If Not isValidProjectName(pName) Then
                                                pName = makeValidProjectName(pName)
                                            End If

                                        End With
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectname in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des ProjektNamens in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try


                                    'Find ProjectTemplate
                                    Dim projTmp As String
                                    Try
                                        Dim projTmpConfig As clsConfigProjectsImport = projectConfig("ProjectTemplate")
                                        If currentWS.Index <> projTmpConfig.sheet And projTmpConfig.sheet <> 0 Then
                                            If Not IsNothing(projTmpConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projTmpConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projTmpConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projTmpConfig.objType = "direkt" Then
                                            vorlagenName = projTmpConfig.content
                                        Else
                                            With currentWS
                                                Select Case projTmpConfig.Typ
                                                    Case "Text"
                                                        projTmp = CStr(.Cells(i, projTmpConfig.column.von).value)
                                                    Case "Integer"
                                                        projTmp = CInt(.Cells(i, projTmpConfig.column.von).value)
                                                    Case "Decimal"
                                                        projTmp = CDbl(.Cells(i, projTmpConfig.column.von).value)
                                                    Case "Date"
                                                        projTmp = CDate(.Cells(i, projTmpConfig.column.von).value)
                                                    Case Else
                                                        projTmp = .Cells(i, projTmpConfig.column.von).value
                                                End Select

                                                If projTmpConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projTmpConfig.content)
                                                    Dim match As Match = regexpression.Match(projTmp)
                                                    If match.Success Then
                                                        projTmp = match.Value
                                                    End If
                                                End If
                                                vorlagenName = projTmp
                                            End With
                                        End If
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the project-template in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des Projekt-Template in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    'Find ProjectStart
                                    Dim projStart As String = ""
                                    Try
                                        Dim projStartConfig As clsConfigProjectsImport = projectConfig("ProjectStart")
                                        If currentWS.Index <> projStartConfig.sheet And projStartConfig.sheet <> 0 Then
                                            If Not IsNothing(projStartConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projStartConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projStartConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projStartConfig.objType = "direkt" Then
                                            startDate = CDate(projStartConfig.content)
                                        Else
                                            With currentWS

                                                Select Case projStartConfig.Typ
                                                    Case "Text"
                                                        projStart = CStr(.Cells(i, projStartConfig.column.von).value)
                                                    Case "Integer"
                                                        projStart = CInt(.Cells(i, projStartConfig.column.von).value)
                                                    Case "Decimal"
                                                        projStart = CDbl(.Cells(i, projStartConfig.column.von).value)
                                                    Case "Date"
                                                        projStart = CDate(currentWS.Cells(i, projStartConfig.column.von).value)
                                                    Case Else
                                                        projStart = .Cells(i, projStartConfig.column.von).value
                                                End Select

                                                If projStartConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projStartConfig.content)
                                                    Dim match As Match = regexpression.Match(projStart)
                                                    If match.Success Then
                                                        projStart = match.Value
                                                    End If
                                                End If
                                            End With
                                            startDate = projStart
                                        End If


                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectstart in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des Projekt-Starts in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    'Find ProjectEnde
                                    Dim projEnde As String = ""
                                    Try
                                        Dim projEndeConfig As clsConfigProjectsImport = projectConfig("ProjectEnd")
                                        If currentWS.Index <> projEndeConfig.sheet And projEndeConfig.sheet <> 0 Then
                                            If Not IsNothing(projEndeConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projEndeConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projEndeConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projEndeConfig.objType = "direkt" Then
                                            endDate = CDate(projEndeConfig.content)
                                        Else

                                            With currentWS
                                                Select Case projEndeConfig.Typ
                                                    Case "Text"
                                                        projEnde = CStr(.Cells(i, projEndeConfig.column.von).value)
                                                    Case "Integer"
                                                        projEnde = CInt(.Cells(i, projEndeConfig.column.von).value)
                                                    Case "Decimal"
                                                        projEnde = CDbl(.Cells(i, projEndeConfig.column.von).value)
                                                    Case "Date"
                                                        projEnde = CDate(.Cells(i, projEndeConfig.column.von).value)
                                                    Case Else
                                                        projEnde = .Cells(i, projEndeConfig.column.von).value
                                                End Select

                                                If projEndeConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projEndeConfig.content)
                                                    Dim match As Match = regexpression.Match(projEnde)
                                                    If match.Success Then
                                                        projEnde = match.Value
                                                    End If
                                                End If
                                            End With
                                            endDate = CDate(projEnde)
                                        End If
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectend in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden des Projekt-Endes in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                        noGo = noGo + 1
                                    End Try

                                    ' find ProjectDescription
                                    Dim projDescr As String = ""
                                    Try
                                        Dim projDescrConfig As clsConfigProjectsImport = projectConfig("ProjectDescription")
                                        If currentWS.Index <> projDescrConfig.sheet And projDescrConfig.sheet <> 0 Then
                                            If Not IsNothing(projDescrConfig.sheet) Then
                                                currentWS = CType(appInstance.Worksheets(projDescrConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            Else
                                                currentWS = CType(appInstance.Worksheets(projDescrConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                            End If
                                        End If

                                        If projDescrConfig.objType = "direkt" Then
                                            description = CStr(projDescrConfig.content)
                                        Else

                                            With currentWS
                                                Select Case projDescrConfig.Typ
                                                    Case "Text"
                                                        projDescr = CStr(.Cells(i, projDescrConfig.column.von).value)
                                                    Case "Integer"
                                                        projDescr = CInt(.Cells(i, projDescrConfig.column.von).value)
                                                    Case "Decimal"
                                                        projDescr = CDbl(.Cells(i, projDescrConfig.column.von).value)
                                                    Case "Date"
                                                        projDescr = CDate(.Cells(i, projDescrConfig.column.von).value)
                                                    Case Else
                                                        projDescr = .Cells(i, projDescrConfig.column.von).value
                                                End Select

                                                If projDescrConfig.objType = "RegEx" Then
                                                    regexpression = New Regex(projDescrConfig.content)
                                                    Dim match As Match = regexpression.Match(projDescr)
                                                    If match.Success Then
                                                        projDescr = match.Value
                                                    End If
                                                End If
                                                description = projDescr
                                            End With
                                        End If
                                    Catch ex As Exception
                                        If awinSettings.englishLanguage Then
                                            outputline = "Couldn't find the projectdescription in line " & i.ToString & "of the inputfile"
                                        Else
                                            outputline = "Fehler beim Herausfinden der Projekt-Beschreibung in Zeile " & i.ToString & " des Inputfiles"
                                        End If
                                        meldungen.Add(outputline)
                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                    End Try



                                    ' Find TimeUnit
                                    Dim timeUnit As String = ""
                                    Dim timeUnitConfig As clsConfigProjectsImport = projectConfig("TimeUnit")
                                    If currentWS.Index <> timeUnitConfig.sheet Then
                                        If Not IsNothing(timeUnitConfig.sheet) Then
                                            currentWS = CType(appInstance.Worksheets(timeUnitConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                        Else
                                            currentWS = CType(appInstance.Worksheets(timeUnitConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                        End If
                                    End If

                                    With currentWS
                                        Select Case timeUnitConfig.Typ
                                            Case "Text"
                                                timeUnit = CStr(.Cells(i, timeUnitConfig.column.von).value)
                                            Case "Integer"
                                                timeUnit = CInt(.Cells(i, timeUnitConfig.column.von).value)
                                            Case "Decimal"
                                                timeUnit = CDbl(.Cells(i, timeUnitConfig.column.von).value)
                                            Case "Date"
                                                timeUnit = CDate(.Cells(i, timeUnitConfig.column.von).value)
                                            Case Else
                                                timeUnit = .Cells(i, timeUnitConfig.column.von).value
                                        End Select

                                        If timeUnitConfig.objType = "RegEx" Then
                                            regexpression = New Regex(timeUnitConfig.content)
                                            Dim timeUnitMatch As Match = regexpression.Match(timeUnit)
                                            Dim xx As MatchCollection = regexpression.Matches(timeUnit)
                                            If timeUnitMatch.Success Then
                                                timeUnit = CStr(timeUnitMatch.Value)
                                                ' find months
                                                Dim months As String = ""
                                                'Dim monthVon As Integer = 0
                                                'Dim monthBis As Integer = 0
                                                Dim monthsConfig As clsConfigProjectsImport = projectConfig("months")
                                                If currentWS.Index <> monthsConfig.sheet Then
                                                    If Not IsNothing(monthsConfig.sheet) Then
                                                        currentWS = CType(appInstance.Worksheets(monthsConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    Else
                                                        currentWS = CType(appInstance.Worksheets(monthsConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    End If
                                                End If
                                                With currentWS
                                                    If Not IsNothing(monthsConfig.Typ) And
                                                        Not IsNothing(monthsConfig.column) Then

                                                        Select Case monthsConfig.Typ
                                                            Case "Text"
                                                                months = CStr(.Cells(i, monthsConfig.column.von).value)
                                                            Case "Integer"
                                                                months = CInt(.Cells(i, monthsConfig.column.von).value)
                                                            Case "Decimal"
                                                                months = CDbl(.Cells(i, monthsConfig.column.von).value)
                                                            Case "Date"
                                                                months = CDate(.Cells(i, monthsConfig.column.von).value)
                                                            Case Else
                                                                months = .Cells(i, monthsConfig.column.von).value
                                                        End Select
                                                    End If

                                                    If Not IsNothing(monthsConfig.cellrange) Then
                                                        If monthsConfig.cellrange Then
                                                            monthVon = monthsConfig.column.von
                                                            monthBis = monthsConfig.column.bis
                                                        End If
                                                    End If

                                                    If Not IsNothing(monthsConfig.objType) Then
                                                        If monthsConfig.objType = "RegEx" Then
                                                            regexpression = New Regex(monthsConfig.content)
                                                            Dim match As Match = regexpression.Match(months)
                                                            If match.Success Then
                                                                months = CInt(match.Value)
                                                            End If
                                                        End If
                                                    End If


                                                End With

                                                ' Find Role
                                                Dim roleName As String
                                                Dim roleNameConfig As clsConfigProjectsImport = projectConfig("Ressourcen")
                                                If currentWS.Index <> roleNameConfig.sheet Then
                                                    If Not IsNothing(roleNameConfig.sheet) Then
                                                        currentWS = CType(appInstance.Worksheets(roleNameConfig.sheet), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    Else
                                                        currentWS = CType(appInstance.Worksheets(roleNameConfig.sheetDescript), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                                    End If
                                                End If

                                                With currentWS
                                                    Select Case roleNameConfig.Typ
                                                        Case "Text"
                                                            roleName = CStr(.Cells(i, roleNameConfig.column.von).value)
                                                        Case "Integer"
                                                            roleName = CInt(.Cells(i, roleNameConfig.column.von).value)
                                                        Case "Decimal"
                                                            roleName = CDbl(.Cells(i, roleNameConfig.column.von).value)
                                                        Case "Date"
                                                            roleName = CDate(.Cells(i, roleNameConfig.column.von).value)
                                                        Case Else
                                                            roleName = .Cells(i, roleNameConfig.column.von).value
                                                    End Select

                                                    If roleNameConfig.objType = "RegEx" Then
                                                        If IsNothing(roleNameConfig.content) Then
                                                            'Fehlermeldung einbauen
                                                            If awinSettings.englishLanguage Then
                                                                outputline = "There is no regular expression defined in the config for getting the rolename"
                                                            Else
                                                                outputline = "Es wurde keine Regular Expression für die Ressource definiert"
                                                            End If
                                                            meldungen.Add(outputline)
                                                            Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                                            noGo = noGo + 1
                                                        Else
                                                            regexpression = New Regex(roleNameConfig.content)
                                                            Dim col As MatchCollection = regexpression.Matches(roleName)
                                                            ' Loop through Matches.
                                                            For Each m As Match In col
                                                                ' Access first Group and its value.
                                                                Dim g As Group = m.Groups(1)
                                                                roleName = g.Value
                                                            Next
                                                        End If

                                                    End If

                                                    If RoleDefinitions.containsName(roleName) Then
                                                        Dim hroleValues As Double()
                                                        ' initialisieren des Array
                                                        ReDim hroleValues(monthBis - monthVon)
                                                        For m = monthVon To monthBis
                                                            Select Case timeUnit
                                                                Case "hours", "hour"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value) / 8
                                                                Case "days", "day"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value)
                                                                Case "weeks", "week"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value) * 5
                                                                Case "months", "month"
                                                                    hroleValues(m - monthVon) = CDbl(.Cells(i, m).value) * nrOfDaysMonth
                                                            End Select

                                                        Next
                                                        If Not roleListNameValues.ContainsKey(roleName) Then
                                                            ' liste aufbauen, die später dazu dient, das erstellte Projekt zu befüllen
                                                            roleListNameValues.Add(roleName, hroleValues)
                                                        Else
                                                            ' evt. aufsummieren der jeweiligen werte eines Monats
                                                        End If
                                                    Else
                                                        If awinSettings.englishLanguage Then
                                                            outputline = "Role " & roleName & " isn't defined in this VC"
                                                        Else
                                                            outputline = "Rolle " & roleName & " existiert in diesem VC nicht"
                                                        End If

                                                        'meldungen.Add(outputline)
                                                        Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)

                                                    End If
                                                End With

                                            End If
                                        End If

                                    End With
                                Next i

                            Catch ex As Exception
                                If awinSettings.englishLanguage Then
                                    outputline = "The actual file isn't conform with the Configuration!"
                                Else
                                    outputline = "die ausgewählte Datei entspricht nicht der Konfiguration!"
                                End If

                                meldungen.Add(outputline)
                                Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                                noGo = noGo + 1
                            End Try

                        End If

                    End If


                    projectWB.Close(SaveChanges:=False)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        outputline = "There is something wrong with the inputfile!"
                    Else
                        outputline = "Fehler im Inputfile!"
                    End If

                    meldungen.Add(outputline)
                    Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
                End Try
            Else
                If awinSettings.englishLanguage Then
                    outputline = "The file you selected doesn't exist!"
                Else
                    outputline = "Die ausgewählte Datei existiert nicht!"
                End If
                Call logfileSchreiben(outputline, "readProjectsWithConfig", anzFehler)
            End If

        Catch ex As Exception

        End Try

        result = (anz_Proj = ImportProjekte.Count) And (anz_Proj > 0)

        readProjectsWithConfig = result
    End Function
End Module
