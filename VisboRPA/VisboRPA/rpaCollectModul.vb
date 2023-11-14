Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports DBAccLayer
Imports WebServerAcc
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
'Imports System.ComponentModel
'Imports System.Windows
'Imports System.Windows.Forms
'Imports System.Security.Principal
Imports System.Text.RegularExpressions
'Imports System.Globalization
Module rpaCollectModul


    ''' <summary>
    ''' liest die im Collect Folder liegende Zeuss Std Datei im Diretory ../collect
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' 
    ''' Ähnlich wie readAvailability... allerdings beschränkzt auf Import Type = 1
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function readRpaAvailabilityOfRoleWithConfig(ByVal kapaConfig As SortedList(Of String, clsConfigKapaImport),
                                                ByVal kapaFileName As String,
                                                ByRef oPCollection As Collection,
                                                ByRef unknownNames As Collection,
                                                ByRef roleMonthList As SortedList(Of String, List(Of Integer)),
                                                ByRef vonDate As Date, ByRef bisDate As Date) As Boolean

        Dim err As New clsErrorCodeMsg
        Dim old_oPCollectionCount As Integer = oPCollection.Count
        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        'Dim fehler As Boolean = False

        Dim spalte As Integer = 2
        Dim firstUrlspalte As Integer = 0
        Dim firstUrlzeile As Integer = 0
        Dim noColor As Integer = -4142
        Dim whiteColor As Integer = 2
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim dateConsidered As Date

        Dim ok2 As Boolean
        Dim isdate As Boolean

        'Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
        Dim monthName As String = ""

        ' tk wird nicht verwendet ... 
        'Dim monthNumber As Integer = 0

        Dim Jahr As Integer = 0
        Dim anzMonthDays As Integer = 0
        Dim colDate As Integer = 0
        Dim anzDays As Integer = 0

        Dim lastZeile As Integer
        'Dim lastSpalte As Integer
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
                ' looking for one zeuss table within File .. 
                Dim relevantTableFound As Boolean = False

                Try
                    For index = 1 To appInstance.Worksheets.Count

                        currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        With currentWS

                            ' tk 7.10.23 wenn ConsideredDate gleich bestimmt werden kann 
                            If Not IsNothing(.Cells(kapaConfig("month").row, kapaConfig("month").column).value) Then
                                Try
                                    dateConsidered = CDate(.Cells(kapaConfig("month").row, kapaConfig("month").column).value)
                                    If DateDiff(DateInterval.Month, StartofCalendar, dateConsidered) > 0 Then
                                        relevantTableFound = True
                                        ok2 = True
                                        isdate = True
                                    Else
                                        ok2 = False
                                        isdate = False
                                    End If

                                Catch ex As Exception
                                    dateConsidered = Date.MinValue
                                    ok2 = False
                                    isdate = False
                                End Try
                            Else
                                ok2 = False
                            End If

                            ' tk 10.10 bei mehreren Blättern führt das zum Fehler ...
                            'If Not ok2 Then

                            '    Dim hjahr As String = CStr(.Cells(kapaConfig("year").row, kapaConfig("year").column).value).Trim
                            '    If IsNothing(hjahr) Then
                            '        Jahr = 0
                            '    Else
                            '        If kapaConfig("year").regex = "RegEx" Then
                            '            'regexpression = New Regex("[0-9]{4}")
                            '            regexpression = New Regex(kapaConfig("year").content)
                            '            Dim match As Match = regexpression.Match(hjahr)
                            '            If match.Success Then
                            '                Jahr = CInt(match.Value)
                            '            Else
                            '                Jahr = 0
                            '            End If
                            '        End If
                            '    End If


                            '    ' Auslesen des relevanten Monats
                            '    Dim hmonth As String = CStr(.Cells(kapaConfig("month").row, kapaConfig("month").column).value).Trim
                            '    If IsNothing(hmonth) Then
                            '        monthName = ""
                            '    Else
                            '        If kapaConfig("month").regex = "RegEx" Then
                            '            regexpression = New Regex(kapaConfig("month").content)
                            '            Dim Match As Match = regexpression.Match(hmonth)
                            '            If Match.Success Then
                            '                monthName = Match.Value
                            '            Else
                            '                monthName = ""
                            '            End If
                            '        End If
                            '    End If

                            '    ' tk 3.2.20 
                            '    isdate = DateTime.TryParse(monthName & " " & Jahr.ToString, dateConsidered)

                            'End If


                            ' Auslesen erste Verfügbarkeitsspalte
                            firstUrlspalte = kapaConfig("valueStart").column
                            firstUrlzeile = kapaConfig("valueStart").row
                        End With

                        If ok2 Then
                            ' weitermachen 
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
                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

                            ElseIf Not isdate Then

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Error in Month of capacity definition: no valid month, year in Capa file: " & kapaFileName
                                Else
                                    msgtxt = "Fehler in Angabe des auszulesenden Monats in Kapazitäts-Datei: " & kapaFileName

                                End If
                                oPCollection.Add(msgtxt)
                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                            Else


                                colDate = getColumnOfDate(dateConsidered)

                                monthDays.Clear()

                                anzMonthDays = DateTime.DaysInMonth(Year(dateConsidered), Month(dateConsidered))
                                'anzMonthDays = DateTime.DaysInMonth(Jahr, Month(dateConsidered))
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
                                    Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

                                Else
                                    ' hier ist sichergestellt, dass die erste Spalte mit 1 beginnt, die letzte Spalte dem Tag entspricht, mit dem der Monat endet

                                    ok = True

                                    anzDays = 0

                                    'lastSpalte = CType(currentWS.Cells(firstUrlzeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                                    lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                                End If

                                If Not ok Then

                                    'fehler = True

                                    If awinSettings.englishLanguage Then
                                        msgtxt = "Error reading Zeuss Information: Please check the file and the according config file definition ..."
                                    Else
                                        msgtxt = "Fehler beim Lesen der Zeuss Informationen: Bitte prüfen Sie die Datei und die Config File Definition ..."
                                    End If
                                    If Not oPCollection.Contains(msgtxt) Then
                                        oPCollection.Add(msgtxt, msgtxt)
                                    End If

                                    Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

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

                                    Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)

                                Else
                                    ' processedNames contains all already considered names; doublettes are recognized and not handeled again ... 
                                    Dim processedNames As New Collection

                                    ' define dateVon and datebis, is needed for input in following methods in calling procedure
                                    If getColumnOfDate(vonDate) > getColumnOfDate(dateConsidered) Then
                                        vonDate = dateConsidered
                                    End If

                                    If getColumnOfDate(bisDate) < getColumnOfDate(dateConsidered) Then
                                        bisDate = dateConsidered
                                    End If

                                    For iZ = firstUrlzeile To lastZeile

                                        If Not IsNothing(CType(currentWS.Cells(iZ, kapaConfig("role").column), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                            rolename = CType(CType(currentWS.Cells(iZ, kapaConfig("role").column), Global.Microsoft.Office.Interop.Excel.Range).Value, String).Trim

                                            Dim checkValue As Double = -1.0
                                            ' it there is something in Spalte AN, then it is a check Value for Testing
                                            If Not IsNothing(CType(currentWS.Cells(iZ, "AN"), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                                Try
                                                    checkValue = CDbl(CType(currentWS.Cells(iZ, "AN"), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                Catch ex As Exception

                                                End Try
                                            End If

                                                ' check whether or not there is a vblf in the name 
                                                If rolename.Contains(vbLf) Then
                                                rolename = rolename.Replace(vbLf, " ").Trim
                                            End If



                                            If rolename <> "" Then
                                                If RoleDefinitions.containsName(rolename) Then
                                                    hrole = RoleDefinitions.getRoledef(rolename)
                                                    If Not IsNothing(hrole) Then

                                                        ' now check whether or not it has already been processed 
                                                        If Not processedNames.Contains(rolename) Then

                                                            processedNames.Add(rolename, rolename)

                                                            Dim defaultHrsPerdayForThisPerson As Double = hrole.defaultDayCapa

                                                            Dim iSp As Integer = firstUrlspalte
                                                            Dim anzArbTage As Double = 0
                                                            Dim anzArbStd As Double = 0

                                                            For Each kvp As KeyValuePair(Of Integer, Integer) In monthDays

                                                                Dim colOfDate As Integer = kvp.Key
                                                                anzDays = kvp.Value

                                                                For sp = iSp + 0 To iSp + anzDays - 1

                                                                    ' tk may lead to neglecting last days, because if there is no entry in the last columns, the lastSpalte is too early
                                                                    'If iSp <= lastSpalte Then

                                                                    Dim hint As Integer = CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex)

                                                                    If CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex) = noColor _
                                                                                Or CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex = whiteColor Then

                                                                        Dim aktCell As Object = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value

                                                                        If Not IsNothing(aktCell) Then

                                                                            If IsNumeric(aktCell) Then

                                                                                Dim angabeInStd As Double = CType(aktCell, Double)

                                                                                If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                                    anzArbStd = anzArbStd + angabeInStd
                                                                                Else
                                                                                    If awinSettings.englishLanguage Then
                                                                                        msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                                    Else
                                                                                        msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                                    End If
                                                                                    If Not oPCollection.Contains(msgtxt) Then
                                                                                        oPCollection.Add(msgtxt, msgtxt)
                                                                                    End If

                                                                                    Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                                                End If
                                                                            Else
                                                                                Dim workHours As String = CType(aktCell, String).Trim
                                                                                If workHours = "" Then
                                                                                    ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                                Else
                                                                                    If kapaConfig("valueSign").regex = "RegEx" Then
                                                                                        regexpression = New Regex(kapaConfig("valueSign").content)
                                                                                        If Not IsNothing(workHours) Then
                                                                                            Dim match As Match = regexpression.Match(workHours)
                                                                                            If match.Success Then
                                                                                                workHours = match.Value
                                                                                                ' Feld ist weiss, oder hat keine Farbe, keine Zahl und keinen "/": also ist es Arbeitstag mit Default-Std pro Tag 
                                                                                                anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                                            End If
                                                                                        End If
                                                                                    End If

                                                                                End If
                                                                            End If

                                                                        Else

                                                                            ' Feld ist ohne Inhalt: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                            anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson


                                                                        End If
                                                                    End If


                                                                Next

                                                                anzArbTage = anzArbStd / 8

                                                                ' now check if there i a checkValue 
                                                                If checkValue >= 0 Then
                                                                    If System.Math.Abs(checkValue - anzArbTage) <= 0.00001 Then
                                                                        msgtxt = "Check correct: Capa = " & anzArbTage & " for " & rolename
                                                                        Call logger(ptErrLevel.logInfo, msgtxt, kapaFileName, anzFehler)
                                                                    Else
                                                                        msgtxt = "Check NOT correct: Check <> Capa  " & checkValue & " <> " & anzArbTage & " for " & rolename
                                                                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                                        oPCollection.Add(msgtxt)
                                                                    End If
                                                                End If

                                                                'nur wenn die hrole schon im Unternehmen ist und nicht wieder ausgetreten ist, wird die Capa eingetragen
                                                                If colOfDate >= getColumnOfDate(hrole.entryDate) And colOfDate < getColumnOfDate(hrole.exitDate) Then
                                                                    hrole.kapazitaet(colOfDate) = anzArbTage
                                                                Else
                                                                    hrole.kapazitaet(colOfDate) = 0
                                                                End If
                                                                iSp = iSp + anzDays
                                                                anzArbTage = 0              ' Anzahl Arbeitstage wieder zurücksetzen für den nächsten Monat
                                                                anzArbStd = 0               ' Anzahl zu leistender Arbeitsstunden wieder zurücksetzen für den nächsten Monat

                                                                ' now remember for the subsequent RPA step of potentially applying %Capa Modifier, that role/month combination was set
                                                                If Not IsNothing(roleMonthList) Then
                                                                    If roleMonthList.ContainsKey(hrole.name) Then
                                                                        If Not roleMonthList.Item(hrole.name).Contains(colOfDate) Then
                                                                            roleMonthList.Item(hrole.name).Add(colOfDate)
                                                                        Else
                                                                            ' do nothing , colOfDate is already in there ...
                                                                        End If
                                                                    Else
                                                                        Dim myList As New List(Of Integer) From {
                                                                            colOfDate
                                                                        }
                                                                        roleMonthList.Add(hrole.name, myList)
                                                                    End If
                                                                End If

                                                            Next


                                                        Else
                                                            ' Error Logging 
                                                            If awinSettings.englishLanguage Then
                                                                msgtxt = "doublette found and exit: " & rolename
                                                            Else
                                                                msgtxt = "Duplikat gefunden, führte zu Abbruch: " & rolename
                                                            End If
                                                            If Not oPCollection.Contains(msgtxt) Then
                                                                oPCollection.Add(msgtxt, msgtxt)
                                                            End If
                                                            Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                                        End If

                                                    Else
                                                        If unknownNames.Contains(rolename) Then
                                                            ' do nothing 
                                                        Else
                                                            unknownNames.Add(rolename, rolename)
                                                        End If
                                                        'If awinSettings.englishLanguage Then
                                                        '    msgtxt = rolename & " not defined ... in VISBO Resource Pool / Organisation"
                                                        'Else
                                                        '    msgtxt = rolename & " nicht definiert im Ressourcen Pool / Organisation"
                                                        'End If

                                                        'Call logger(ptErrLevel.logWarning, msgtxt, kapaFileName, anzFehler)
                                                    End If
                                                End If

                                            Else

                                                If awinSettings.englishLanguage Then
                                                    msgtxt = "No Name given ..."
                                                Else
                                                    msgtxt = "kein Name angegeben ..."
                                                End If
                                                If Not oPCollection.Contains(msgtxt) Then
                                                    oPCollection.Add(msgtxt, msgtxt)
                                                End If
                                                Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                                            End If

                                        End If

                                    Next iZ

                                End If   ' ende von if not OK


                            End If      'beginningDay = 1
                        Else
                            ' no Error Messages : check whether or no a relevant Table was found 
                            'If awinSettings.englishLanguage Then
                            '    msgtxt = "Error in date of Month: " & dateConsidered.ToShortDateString
                            'Else
                            '    msgtxt = "Fehler im Datum: " & dateConsidered.ToShortDateString

                            'End If
                            'oPCollection.Add(msgtxt)
                            'Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                        End If


                    Next index

                    If Not relevantTableFound Then
                        If awinSettings.englishLanguage Then
                            msgtxt = "nothing to read in " & kapaFileName
                        Else
                            msgtxt = "keine Daten zu lesen in " & kapaFileName

                        End If
                        oPCollection.Add(msgtxt)
                        Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                    End If


                Catch ex2 As Exception
                    If awinSettings.englishLanguage Then
                        msgtxt = "Error reading dates like month/year in " & kapaFileName & vbLf & ex2.Message
                    Else
                        msgtxt = "Fehler beim Lesen der notwendigen Randdaten wie Monat/Jahr in " & kapaFileName & vbLf & ex2.Message
                    End If
                    If Not oPCollection.Contains(msgtxt) Then
                        oPCollection.Add(msgtxt, msgtxt)
                    End If
                    Call logger(ptErrLevel.logError, msgtxt, kapaFileName, anzFehler)
                End Try

                kapaWB.Close(SaveChanges:=False)

            Catch ex As Exception

                kapaWB.Close(SaveChanges:=False)
            End Try

        End If



        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True


        readRpaAvailabilityOfRoleWithConfig = (oPCollection.Count = old_oPCollectionCount)

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dateiName"></param>
    ''' <param name="meldungen"></param>
    ''' <param name="roleMonthList"></param>
    ''' <param name="applyPercent"></param>
    ''' <param name="namesProcessed"></param>
    ''' <returns></returns>
    Friend Function readRpaKapaModifier(ByVal dateiName As String, ByRef meldungen As Collection,
                                        ByVal roleMonthList As SortedList(Of String, List(Of Integer)), ByVal applyPercent As Boolean,
                                        ByVal namesProcessed As SortedList(Of String, String)) As Boolean


        Dim ok As Boolean = True

        Dim endeZeile As Integer

        Dim endeZeile1 As Integer
        Dim endeZeile2 As Integer
        Dim spalte As Integer = 2
        Dim blattname As String = "Kapazität"
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date
        Dim tmpKapa As Double
        Dim lastSpalte As Integer
        Dim errMsg As String = ""

        Dim noError As Boolean = True

        Dim colName As Integer = 2
        Dim colPersNr As Integer = 1
        endeZeile = 0

        If Not IsNothing(dateiName) Then

            If My.Computer.FileSystem.FileExists(dateiName) And dateiName.Contains("Kapazität") And dateiName.Contains("Modifier") Then

                Try
                    appInstance.Workbooks.Open(dateiName)
                    ok = True

                    Try

                        currentWS = CType(appInstance.Worksheets(blattname), Global.Microsoft.Office.Interop.Excel.Worksheet)

                        Try
                            endeZeile1 = CType(currentWS.Cells(12000, "A"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row + 1
                            endeZeile2 = CType(currentWS.Cells(12000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row + 1
                            endeZeile = System.Math.Max(endeZeile1, endeZeile2)
                        Catch ex As Exception
                            endeZeile = 0
                        End Try


                        If endeZeile > 2 Then

                            lastSpalte = CType(currentWS.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column

                            ' bestimme jetzt die Spalte, wo der Name stehen sollte 
                            Dim dateFound As Boolean = False
                            Dim tmpSpalte As Integer = 2
                            Do While Not dateFound
                                If Not IsNothing(CType(currentWS.Cells(1, tmpSpalte), Excel.Range).Value) Then
                                    If IsDate(CType(currentWS.Cells(1, tmpSpalte), Excel.Range).Value) Then
                                        dateFound = True
                                        colName = tmpSpalte - 1
                                        colPersNr = tmpSpalte - 2
                                    Else
                                        tmpSpalte = tmpSpalte + 1
                                    End If
                                End If
                            Loop

                            ' jetzt wird Zeile für Zeile nachgesehen, ob das eine Basic Role ist und dann die Kapas besetzt 

                            Dim aktzeile As Integer = 2
                            Do While aktzeile < endeZeile

                                Dim subRole As clsRollenDefinition = Nothing

                                Dim subRolePersNr As String = CStr(CType(currentWS.Cells(aktzeile, colPersNr), Excel.Range).Value)
                                If Not IsNothing(subRolePersNr) Then
                                    subRolePersNr = subRolePersNr.Trim
                                    If subRolePersNr <> "" Then
                                        subRole = RoleDefinitions.getRoledefByEmployeeNr(subRolePersNr)
                                    End If
                                End If

                                Dim subRoleName As String = CStr(CType(currentWS.Cells(aktzeile, colName), Excel.Range).Value)
                                If Not IsNothing(subRoleName) Then
                                    subRoleName = subRoleName.Trim
                                    If IsNothing(subRole) Then
                                        If subRoleName.Length > 0 Then
                                            If RoleDefinitions.containsName(subRoleName) Then
                                                subRole = RoleDefinitions.getRoledef(subRoleName)
                                            End If
                                        End If

                                    Else
                                        ' do the check 
                                        Dim checkRole As clsRollenDefinition = Nothing
                                        If subRoleName.Length > 0 Then
                                            If RoleDefinitions.containsName(subRoleName) Then
                                                checkRole = RoleDefinitions.getRoledef(subRoleName)
                                                If checkRole.name <> subRole.name Then
                                                    ' Protocol it , that nr and according name do not mactch to each other 
                                                    msgTxt = "Personal number and Name does not match: " & subRolePersNr & " <> " & subRoleName & "continued with Employee-Nr referencing -> " & subRole.name
                                                    Call logger(ptErrLevel.logWarning, msgTxt, dateiName, anzFehler)
                                                End If
                                            End If
                                        End If

                                    End If
                                End If


                                If Not IsNothing(subRole) Then


                                    ' nur weiter machen, wenn es keine SummenRolle ist ...
                                    If Not subRole.isCombinedRole Then

                                        If Not namesProcessed.ContainsKey(subRole.name) Then

                                            namesProcessed.Add(subRole.name, dateiName)

                                            Try
                                                spalte = colName + 1
                                                tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                                                ' erstmal dahin positionieren, wo das Datum auch mit oder nach StartOfCalendar beginnt  

                                                Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) < 0 And spalte <= lastSpalte
                                                    Try
                                                        spalte = spalte + 1
                                                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                                    Catch ex As Exception

                                                    End Try
                                                Loop

                                                Do While spalte < 241 And spalte <= lastSpalte

                                                    Try
                                                        index = getColumnOfDate(tmpDate)
                                                        If index >= 1 Then
                                                            tmpKapa = CDbl(CType(currentWS.Cells(aktzeile, spalte), Excel.Range).Value)

                                                            If tmpKapa >= 0 Then
                                                                ' allow only valid values ge 0 
                                                                Dim myDisplayFormat As String = CType(CType(currentWS.Cells(aktzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).DisplayFormat, Excel.DisplayFormat).NumberFormat
                                                                Dim cellISPercent As Boolean = myDisplayFormat.Contains("%")

                                                                If index <= 240 And index > 0 And tmpKapa >= 0 Then
                                                                    If index >= getColumnOfDate(subRole.entryDate) And index < getColumnOfDate(subRole.exitDate) Then
                                                                        If applyPercent And cellISPercent Then
                                                                            If tmpKapa <= 1.0 And Not IsNothing(roleMonthList) Then
                                                                                ' prozentuale Anwendung ... aber nur, wenn zuvor über z.B Zeuss Import ein absoluter Wert gelesen wurde 
                                                                                ' andernfalls wäre eine wiederholte Anwendung von applyPercent möglich -> nicht die Absicht !
                                                                                If roleMonthList.ContainsKey(subRole.name) Then
                                                                                    If roleMonthList.Item(subRole.name).Contains(index) Then
                                                                                        subRole.kapazitaet(index) = subRole.kapazitaet(index) * tmpKapa
                                                                                    Else
                                                                                        msgTxt = "capacity remains unchanged: no data provided for role " & subRole.name & " month " & getDateofColumn(index, False).ToString(“MM-yy”, Globalization.CultureInfo.InvariantCulture) & " in previous import step"
                                                                                        Call logger(ptErrLevel.logWarning, msgTxt, dateiName, anzFehler)
                                                                                    End If
                                                                                Else
                                                                                    msgTxt = "capacity remains unchanged: no data provided for role " & subRole.name & " at all in previous import step"
                                                                                    Call logger(ptErrLevel.logWarning, msgTxt, dateiName, anzFehler)
                                                                                End If
                                                                            Else
                                                                                msgTxt = "capacity remains unchanged: invalid percentage provided or no data imported previously "
                                                                                meldungen.Add(msgTxt)
                                                                                Call logger(ptErrLevel.logError, msgTxt, dateiName, anzFehler)
                                                                            End If
                                                                        Else
                                                                            subRole.kapazitaet(index) = tmpKapa
                                                                        End If

                                                                    Else
                                                                        subRole.kapazitaet(index) = 0
                                                                    End If

                                                                End If
                                                            End If

                                                        End If

                                                        spalte = spalte + 1
                                                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                                    Catch ex As Exception
                                                        noError = False
                                                        errMsg = ex.Message & vbLf & "File " & dateiName & ": error when setting value for " & subRoleName & " in row, column: " & aktzeile & ", " & spalte
                                                        meldungen.Add(errMsg)
                                                        Call logger(ptErrLevel.logError, errMsg, dateiName, anzFehler)
                                                    End Try


                                                Loop

                                            Catch ex As Exception

                                            End Try
                                        Else
                                            noError = True
                                            errMsg = "Name " & subRole.name & " in " & dateiName & " was ignored  because it has already been processed in " & namesProcessed.Item(subRole.name)

                                            Call logger(ptErrLevel.logWarning, errMsg, dateiName, anzFehler)
                                        End If

                                    Else
                                        noError = True
                                        errMsg = "Name " & subRoleName & " is combinedRole; combinedRoles are calculated automatically" & " (File " & dateiName & " )"

                                        Call logger(ptErrLevel.logWarning, errMsg, dateiName, anzFehler)
                                    End If

                                Else
                                    ' die Rolle existiert überhaupt nicht im Ressourcen Pool 
                                    noError = False
                                    errMsg = "Nr resp. Name " & subRolePersNr & " : " & subRoleName & " does not exist in VISBO Organisation"
                                    Call logger(ptErrLevel.logError, errMsg, dateiName, anzFehler)
                                End If

                                aktzeile = aktzeile + 1
                                ' jetzt spalte wieder auf 2 setzen 
                                spalte = 2
                            Loop

                        Else
                            noError = False
                            errMsg = "File " & dateiName & " does not contain data in column A ..."
                            meldungen.Add(errMsg)
                            Call logger(ptErrLevel.logError, errMsg, dateiName, anzFehler)
                        End If

                    Catch ex2 As Exception
                        noError = False
                        errMsg = "File " & dateiName & ": unidentified error ... "
                        meldungen.Add(errMsg)
                        Call logger(ptErrLevel.logError, errMsg, dateiName, anzFehler)
                    End Try


                    appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Catch ex As Exception
                    appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                End Try

            Else
                noError = False
                errMsg = "File " & dateiName & ": name has to contain strings 'Modifier' and 'Kapazität'"
                meldungen.Add(errMsg)
                Call logger(ptErrLevel.logError, errMsg, dateiName, anzFehler)
            End If

        End If

        readRpaKapaModifier = noError
    End Function


End Module
