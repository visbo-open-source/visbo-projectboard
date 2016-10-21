Imports Microsoft.Office.Interop.Excel
Public Class clsLanguages

    Private _languageItems As SortedList(Of String, Collection)



    ''' <summary>
    ''' gibt zurück, wieviele Sprachen enthalten sind 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count() As Integer
        Get
            count = _languageItems.Count
        End Get
    End Property
    ''' <summary>
    ''' liest die Sprache aus einer übergebenen Excel-Datei (path + Datei-NAme) ein 
    ''' wirft Exception, wenn es nicht klappt
    ''' die erste Spalte (=Original) muss bereits existieren und 100% übereinstimmen, sonst Exception 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub importLanguages()

        Dim userHome As String = My.Computer.FileSystem.SpecialDirectories.Desktop
        Dim excelApp As Excel.Application = Nothing
        Dim excelDidExist As Boolean = False
        Dim fileName As String = userHome & "\" & "PPTlanguages.xlsx"
        Dim ok As Boolean = True

        Try
            ' prüft, ob Excel bereits geöffnet ist
            excelApp = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            excelDidExist = True
        Catch ex As Exception
            ' wenn nein: öffnet Excel 
            Try
                excelApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                excelDidExist = False
            Catch ex1 As Exception
                Call MsgBox("Excel kann nicht gestartet werden ...")
                Exit Sub
            End Try

        End Try

        If My.Computer.FileSystem.FileExists(fileName) Then
            Try
                excelApp.Workbooks.Open(fileName)
            Catch ex As Exception
                Call MsgBox("neues Workbook kann nicht geöffnet werden  ...")
                Exit Sub
            End Try

        Else

            Call MsgBox("File existiert nicht: " & vbLf & fileName)
            Exit Sub

        End If

        ' jetzt werden die Language Files ausgelesen ... 
        Try

            Dim anzSpalten As Integer = 0
            Dim anzZeilen As Integer = 0
            With CType(excelApp.ActiveSheet, Excel.Worksheet)

                anzSpalten = .UsedRange.Columns.Count
                anzZeilen = .UsedRange.Rows.Count

                Dim tmpName As String



                For ixSP As Integer = 1 To anzSpalten
                    tmpName = CStr(CType(.Cells(1, ixSP), Excel.Range).Value)
                    Dim tmpCollection As New Collection

                    For ixZE = 2 To anzZeilen
                        Dim tmpItem As String = CStr(CType(.Cells(ixZE, ixSP), Excel.Range).Value)
                        If IsNothing(tmpItem) Then
                            tmpItem = ""
                        End If
                        If Not tmpCollection.Contains(tmpItem) And tmpItem.Length > 0 Then
                            tmpCollection.Add(tmpItem, tmpItem)
                        End If
                    Next

                    ' jetzt , bei ixSP = 1 prüfen, ob die Original Werte genau identisch sind; 
                    ' andernfalls käme nur Schmarrn raus 
                    If ixSP = 1 Then
                        If tmpName = "Original" Then
                            Dim pruefCollection = _languageItems.Item(tmpName)
                            If pruefCollection.Count = tmpCollection.Count Then
                                Dim k As Integer = 1
                                While k <= tmpCollection.Count And ok
                                    If CStr(pruefCollection.Item(k)) = CStr(tmpCollection.Item(k)) Then
                                        k = k + 1
                                    Else
                                        ok = False
                                    End If
                                End While
                            Else
                                ok = False
                            End If

                        Else
                            ok = False
                        End If
                    End If

                    If Not ok Then
                        Exit For
                    End If

                    ' jetzt die Language hinzufügen
                    If ixSP > 1 Then
                        If Not _languageItems.ContainsKey(tmpName) Then
                            _languageItems.Add(tmpName, tmpCollection)
                        Else
                            _languageItems.Remove(tmpName)
                            _languageItems.Add(tmpName, tmpCollection)
                        End If
                    End If
                    

                Next

            End With

            With CType(excelApp.ActiveWorkbook, Excel.Workbook)
                .Close(SaveChanges:=False)
            End With

        Catch ex As Exception

        End Try

        ' Excel beenden, wenn es nicht vorher bereits existierte ...
        If Not excelDidExist Then
            excelApp.Quit()
        End If

        If Not OK Then
            Throw New Exception("Übersetzungen passen nicht ...")
        End If


    End Sub

    
    ''' <summary>
    ''' exportiert die Sprach Bezeichner in eine Excel Datei mit dem angegebenen Datei-Namen 
    ''' wirft Exception, wenn es nicht klappt
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub exportLanguages()
        Dim userHome As String = My.Computer.FileSystem.SpecialDirectories.Desktop
        Dim excelApp As Excel.Application = Nothing
        Dim excelDidExist As Boolean = False
        Dim fileName As String = userHome & "\" & "PPTlanguages.xlsx"

        Try
            ' prüft, ob Excel bereits geöffnet ist
            excelApp = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            excelDidExist = True
        Catch ex As Exception
            ' wenn nein: öffnet Excel 
            Try
                excelApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                excelDidExist = False
            Catch ex1 As Exception
                Call MsgBox("Excel kann nicht gestartet werden ...")
                Exit Sub
            End Try

        End Try


        If My.Computer.FileSystem.FileExists(fileName) Then
            Try
                excelApp.Workbooks.Open(fileName)
                With CType(excelApp.ActiveSheet, Excel.Worksheet)
                    .UsedRange.Clear()
                End With
            Catch ex As Exception
                Call MsgBox("neues Workbook kann nicht geöffnet werden  ...")
                Exit Sub
            End Try

        Else
            Try
                excelApp.Workbooks.Add()
            Catch ex1 As Exception
                Call MsgBox("neues Workbook kann nicht erstellt werden  ...")
                Exit Sub
            End Try
        End If

        ' das File ist jetzt geöffnet bzw erzeugt  ...  
        Try

            With CType(excelApp.ActiveSheet, Excel.Worksheet)

                Dim spalte As Integer = 1

                For Each kvp As KeyValuePair(Of String, Collection) In _languageItems
                    Dim zeile As Integer = 1
                    Dim lName As String = kvp.Key
                    CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Key

                    zeile = 2
                    For Each item As String In kvp.Value
                        CType(.Cells(zeile, spalte), Excel.Range).Value = item
                        zeile = zeile + 1
                    Next

                    spalte = spalte + 1
                Next
            End With

            With CType(excelApp.ActiveWorkbook, Excel.Workbook)
                .Close(SaveChanges:=True, Filename:=fileName)
            End With

        Catch ex As Exception

        End Try

        ' Excel beenden, wenn es nicht vorher bereits existierte ...
        If Not excelDidExist Then
            excelApp.Quit()
        End If

    End Sub

    ''' <summary>
    ''' fügt eine Sprache mit den Bezeichnern für Phasen / Meilensteine hinzu
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="items"></param>
    ''' <remarks></remarks>
    Public Sub addLanguage(ByVal key As String, ByVal items As Collection)
        If _languageItems.ContainsKey(key) Then
            _languageItems.Item(key) = items
        Else
            _languageItems.Add(key, items)
        End If
    End Sub

    Public Sub New()
        _languageItems = New SortedList(Of String, Collection)
    End Sub


End Class
