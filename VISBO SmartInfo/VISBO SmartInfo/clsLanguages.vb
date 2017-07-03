Imports Microsoft.Office.Interop.Excel
Imports System.Xml
Imports System.Xml.Schema
<Serializable()> Public Class clsLanguages

    Private _languageItems As SortedList(Of String, List(Of String))


    ''' <summary>
    ''' gibt die Sprache mit lfd Nr index zurück
    ''' Index läuft von 1 .. count
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getLanguageName(ByVal index As Integer) As String
        Get
            If index >= 1 And index <= _languageItems.Count Then
                getLanguageName = _languageItems.ElementAt(index - 1).Key
            Else
                getLanguageName = ""
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt die Sprachen-Items zurück, die zur angegebenen Sprache gehören
    ''' </summary>
    ''' <param name="lName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getLanguage(ByVal lName As String) As List(Of String)
        Get
            Dim tmpList As List(Of String) = Nothing
            If _languageItems.ContainsKey(lName) Then
                tmpList = _languageItems.Item(lName)
            Else
                tmpList = Nothing
            End If
            getLanguage = tmpList
        End Get
    End Property

    ''' <summary>
    ''' übersetzt den String anhand der languageTabellen in die gewählte Sprache
    ''' wenn ein isCombinedName angegeben wird, wird der in seine Bestandteile zerlegt, überstezt und wieder zusammengesetzt ... 
    ''' </summary>
    ''' <param name="tmpText"></param>
    ''' <param name="selectedLanguage"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property translate(ByVal tmpText As String, ByVal selectedLanguage As String, _
                                       Optional ByVal elemName As String = "", Optional ByVal isCombinedName As Boolean = False) As String
        Get
            Dim newText As String = tmpText

            If isCombinedName And elemName.Length > 0 Then
                Dim newLength As Integer = tmpText.Length - (elemName.Length + 1)
                Dim tmpText1 As String = tmpText.Substring(0, newLength)
                Dim tmpText2 As String = elemName
                newText = translate(tmpText1, selectedLanguage) & "-" & translate(tmpText2, selectedLanguage)
            Else
                If _languageItems.ContainsKey(selectedLanguage) Then

                    Dim origItems As List(Of String) = _languageItems.Item(defaultSprache)
                    If origItems.Contains(tmpText) Then
                        Dim found As Boolean = False
                        Dim index As Integer = 0
                        Do While index <= origItems.Count - 1 And Not found
                            If CStr(origItems.Item(index)) = tmpText Then
                                found = True
                            Else
                                index = index + 1
                            End If
                        Loop
                        If found Then
                            Dim newLangItems As List(Of String) = _languageItems.Item(selectedLanguage)
                            newText = CStr(newLangItems.Item(index))
                        End If
                    End If
                End If
            End If
            

            translate = newText
        End Get
    End Property

    ''' <summary>
    ''' übersetzt den String anhand der languageTabellen in die Original Sprache
    ''' </summary>
    ''' <param name="tmpText"></param>
    ''' <param name="selectedLanguage"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property backtranslate(ByVal tmpText As String, ByVal selectedLanguage As String) As String
        Get
            Dim newText As String = tmpText

            If _languageItems.ContainsKey(selectedLanguage) Then

                Dim oldLangItems As List(Of String) = _languageItems.Item(selectedLanguage)
                If oldLangItems.Contains(tmpText) Then
                    Dim found As Boolean = False
                    Dim index As Integer = 0
                    Do While index <= oldLangItems.Count - 1 And Not found
                        If CStr(oldLangItems.Item(index)) = tmpText Then
                            found = True
                        Else
                            index = index + 1
                        End If
                    Loop
                    If found Then
                        Dim origItems As List(Of String) = _languageItems.Item(defaultSprache)
                        newText = CStr(origItems.Item(index))
                    End If
                End If
            End If

            backtranslate = newText
        End Get
    End Property


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
        Dim reason As String = ""

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
                    Dim tmpList As New List(Of String)

                    For ixZE = 2 To anzZeilen
                        Dim tmpItem As String = CStr(CType(.Cells(ixZE, ixSP), Excel.Range).Value)
                        If IsNothing(tmpItem) Then
                            tmpItem = ""
                        End If
                        If Not tmpList.Contains(tmpItem) And tmpItem.Length > 0 Then
                            tmpList.Add(tmpItem)
                        End If
                    Next

                    ' jetzt , bei ixSP = 1 prüfen, ob die Original Werte genau identisch sind; 
                    ' andernfalls käme nur Schmarrn raus 
                    If ixSP = 1 Then
                        If tmpName = defaultSprache Then
                            Dim pruefListe As List(Of String) = _languageItems.Item(defaultSprache)
                            If pruefListe.Count <= tmpList.Count Then

                                For Each pruefItem As String In pruefListe
                                    If Not tmpList.Contains(pruefItem) Then
                                        ok = False
                                        reason = "nicht alle Elemente sind in Übersetzungstabelle enthalten"
                                        Exit For
                                    End If
                                Next

                            Else
                                reason = "nicht alle Elemente sind in Übersetzungstabelle enthalten"
                                ok = False
                            End If

                        Else
                            reason = "1. Sprache nicht Default Sprache: Original"
                            ok = False
                        End If
                    End If

                    If Not ok Then
                        Exit For
                    End If

                    ' jetzt die Language hinzufügen, aber nur wenn die Anzahl genauso groß ist wie bei Default-Language 
                    If ixSP > 1 Then
                        If _languageItems.Item(defaultSprache).Count = tmpList.Count Then
                            If Not _languageItems.ContainsKey(tmpName) Then
                                _languageItems.Add(tmpName, tmpList)
                            Else
                                _languageItems.Remove(tmpName)
                                _languageItems.Add(tmpName, tmpList)
                            End If
                        Else
                            reason = "x. Sprache hat nicht identisch viele Einträge wie die Original-Sprache"
                            ok = False
                            Exit For
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

        If Not ok Then
            Throw New Exception("Fehler bei Import: " & vbLf & reason)
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

                For Each kvp As KeyValuePair(Of String, List(Of String)) In _languageItems
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
    Public Sub addLanguage(ByVal key As String, ByVal items As List(Of String))
        If _languageItems.ContainsKey(key) Then
            _languageItems.Item(key) = items
        Else
            _languageItems.Add(key, items)
        End If
    End Sub

    Public Sub New()
        _languageItems = New SortedList(Of String, List(Of String))
    End Sub

    ' ''' <summary>
    ' ''' wird benötigt, um eine XML Struktur aufzubauen, die einen eindimensionalen Array hat und sonst auch nur einfache Datentypen 
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Public Class languageArray

    '    Friend sprachArray As String()

    '    Friend dimen1 As Integer
    '    Friend dimen2 As Integer

    '    Sub New(ByVal dm1 As Integer, ByVal dm2 As Integer)
    '        dimen1 = dm1
    '        dimen2 = dm2
    '        ReDim sprachArray(dm1 * dm2 - 1)
    '    End Sub

    'End Class

End Class
