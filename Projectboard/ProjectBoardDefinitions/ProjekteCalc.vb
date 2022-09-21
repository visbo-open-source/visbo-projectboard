Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports core = Microsoft.Office.Core
Imports xlNS = Microsoft.Office.Interop.Excel
'Imports pptNS = Microsoft.Office.Interop.PowerPoint
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Constants
Imports ProjectBoardDefinitions
Imports System.Globalization


Public Module ProjekteCalc


    Public Function berechneOptimierungsWert(ByRef currentProjektListe As clsProjekte, ByRef DiagrammTyp As String, ByRef myCollection As Collection) As Double
        Dim value As Double
        Dim kennzahl1 As Double
        Dim kennzahl2 As Double
        Dim avgValue As Double

        If DiagrammTyp = DiagrammTypen(1) Then
            value = currentProjektListe.getbadCostOfRole(myCollection)

        ElseIf DiagrammTyp = DiagrammTypen(0) Then

            kennzahl1 = currentProjektListe.getAverage(myCollection, DiagrammTyp)
            kennzahl2 = currentProjektListe.getPhaseSchwellWerteInMonth(myCollection).Sum
            avgValue = System.Math.Max(kennzahl1, kennzahl2)
            value = currentProjektListe.getDeviationfromAverage(myCollection, avgValue, DiagrammTyp)

        ElseIf DiagrammTyp = DiagrammTypen(2) Then
            avgValue = currentProjektListe.getAverage(myCollection, DiagrammTyp)
            value = currentProjektListe.getDeviationfromAverage(myCollection, avgValue, DiagrammTyp)

        ElseIf DiagrammTyp = DiagrammTypen(4) Then
            ' da der Optimierungs-Algorithmus die kleinste Zahl sucht , muss mit -1 multipliziert werden, 
            ' damit tatsächlich der größte Ertrag heraus kommt 
            value = currentProjektListe.getErgebniskennzahl * (-1)

        ElseIf DiagrammTyp = DiagrammTypen(5) Then
            'Throw New ArgumentException("Optimierung ist für diesen Diagramm-Typ nicht implementiert")
            ' tk: das folgende kann aktiviert werden, sobald 
            kennzahl1 = currentProjektListe.getAverage(myCollection, DiagrammTyp)
            kennzahl2 = currentProjektListe.getMilestoneSchwellWerteInMonth(myCollection).Sum
            avgValue = System.Math.Max(kennzahl1, kennzahl2)
            value = currentProjektListe.getDeviationfromAverage(myCollection, avgValue, DiagrammTyp)

        ElseIf DiagrammTyp = DiagrammTypen(9) Then
            'CashFlow 
            value = currentProjektListe.getCashFlow().Sum
        Else
            Throw New ArgumentException("Optimierung ist für diesen Diagramm-Typ nicht implementiert")
        End If

        berechneOptimierungsWert = value

    End Function
    ''' <summary>
    ''' gibt den Schnittmengen-Array zurück, der Array tmpValues hat die Dimension pEnde-PStart
    ''' und stellt die Werte dar, die im Monat pStart .. PEnde gelten. 
    ''' Im Schnittmengen Array sind die Werte der Dimension bis-von
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="pStart"></param>
    ''' <param name="pEnde"></param>
    ''' <param name="tmpValues"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcArrayIntersection(ByVal von As Integer, ByVal bis As Integer,
                                              ByVal pStart As Integer, ByVal pEnde As Integer,
                                              ByVal tmpValues As Double()) As Double()
        Dim intersectionArray() As Double
        ReDim intersectionArray(bis - von)
        Dim startIX As Integer, endIX As Integer
        Dim sonderfall As Boolean = False
        Dim noOverlap As Boolean = False

        If pStart > bis Or von > pEnde Then
            ' es gibt überhaupt keine Überlappung 
            noOverlap = True

        ElseIf von = pStart And bis = pEnde Then
            startIX = 1
            endIX = tmpValues.Length

        ElseIf von <= pStart Then
            startIX = pStart - von + 1

            If bis <= pEnde Then
                endIX = startIX + bis - pStart
            Else
                ' bis ist größer als pEnde, es gilt das Gleiche wie oben  
                endIX = startIX + pEnde - pStart
            End If

            sonderfall = True

        Else
            ' von ist größer als pStart 
            startIX = von - pStart + 1

            If bis <= pEnde Then
                endIX = startIX + bis - von
            Else
                endIX = startIX + pEnde - von
            End If

        End If

        If noOverlap Then
            ' nichts tun 
        Else
            If sonderfall Then
                For i = startIX To endIX
                    intersectionArray(i - 1) = tmpValues(i - startIX)
                Next
            Else
                For i As Integer = startIX To endIX
                    intersectionArray(i - startIX) = tmpValues(i - 1)
                Next
            End If
        End If

        calcArrayIntersection = intersectionArray

    End Function

    Public Function calcBestCandidates(ByVal priorityPeople As SortedList(Of String, Double),
                                       ByVal myCurrentskillID As Integer,
                                       ByVal candidates As SortedList(Of Double, Integer),
                                       ByVal projectScopeCandidates As SortedList(Of Double, Integer),
                                       ByVal valueToSubstitute As Double) As SortedList(Of Double, Integer)

        Dim result As New SortedList(Of Double, Integer)

        Dim prioPeopleIDs As New List(Of Integer)

        ' are priority candidates in the candidates list? 
        For Each prioPerson As KeyValuePair(Of String, Double) In priorityPeople

            Dim prioPersonSkillID As Integer = -1
            Dim prioPersonID As Integer = RoleDefinitions.parseRoleNameID(prioPerson.Key, prioPersonSkillID)

            If Not prioPeopleIDs.Contains(prioPersonID) Then
                prioPeopleIDs.Add(prioPersonID)
            End If

            Dim found As Boolean = False
            Dim i As Integer = candidates.Count - 1

            Do While Not found And i >= 0
                found = (candidates.ElementAt(i).Value = prioPersonID)
                If Not found Then
                    i = i - 1
                End If
            Loop


            If found And valueToSubstitute > 0 Then

                Dim myValue As Double = System.Math.Min(candidates.ElementAt(i).Key, valueToSubstitute)

                If myValue < 0 Then
                    myValue = 0
                End If

                valueToSubstitute = valueToSubstitute - myValue

                If myValue > 0 Then
                    result.Add(myValue, prioPersonID)
                End If

            End If

        Next

        ' now do the rest ..
        If valueToSubstitute > 0 Then
            Dim ix As Integer = candidates.Count - 1
            ' greatest value is at the end
            Do While ix >= 0 And valueToSubstitute > 0

                Dim candidate As KeyValuePair(Of Double, Integer) = candidates.ElementAt(ix)
                If Not prioPeopleIDs.Contains(candidate.Value) Then
                    If candidate.Key >= valueToSubstitute Then
                        If ((result.Count > 0) Or (isAmongTopGroup(projectScopeCandidates, candidate.Value))) Then
                            result.Add(valueToSubstitute, candidate.Value)
                            valueToSubstitute = 0
                        Else
                            Dim myValue As Double = System.Math.Truncate(0.6 * valueToSubstitute)
                            result.Add(myValue, candidate.Value)
                            valueToSubstitute = valueToSubstitute - myValue
                        End If
                    Else
                        result.Add(candidate.Key, candidate.Value)
                        valueToSubstitute = valueToSubstitute - candidate.Key
                    End If
                End If
                ix = ix - 1
            Loop

        End If

        calcBestCandidates = result
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="arrayZahlen"></param>
    ''' <param name="ixLeft"></param>
    ''' <param name="ixRight"></param>
    ''' <returns></returns>
    Public Function calcUtilizationSortCriteria(ByVal arrayZahlen() As Double, ByVal ixLeft As Integer, ByVal ixRight As Integer, ByVal kapaArray() As Double) As Double
        Dim tmpResult As Double = 0.0

        Dim dimension1 As Integer = arrayZahlen.Length - 1
        Dim dimension2 As Integer = kapaArray.Length - 1

        If dimension1 <> dimension2 Then
            ' nichts tun 
        Else
            If ixLeft >= 0 And ixRight <= dimension1 Then
                For i As Integer = ixLeft To ixRight

                    If arrayZahlen(i) > 0 Then
                        If kapaArray(i) > 0 Then
                            tmpResult = tmpResult + System.Math.Pow((arrayZahlen(i) + kapaArray(i)) / kapaArray(i), 3)
                        Else
                            tmpResult = tmpResult + System.Math.Pow(arrayZahlen(i), 2)
                        End If

                    End If

                Next
            End If

        End If



        calcUtilizationSortCriteria = tmpResult
    End Function
    Public Function calcChartKennung(ByVal typ As String, ByVal index As Integer, ByVal mycollection As Collection) As String
        Dim IDkennung As String
        Dim cName As String = ""
        Dim breadcrumb As String = ""

        IDkennung = typ & "#" & index.ToString

        If typ = "pf" Then


            Try
                Select Case index
                    Case PTpfdk.Phasen

                        If mycollection.Count = PhaseDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count

                                cName = splitHryFullnameTo1(CStr(mycollection.Item(i)))
                                'cName = CStr(mycollection.Item(i)).Replace("#", "-")
                                ' der evtl vorhandenen Breadcrumb hat als Trennzeichen das #
                                Try
                                    IDkennung = IDkennung & "#" & cName
                                Catch ex As Exception
                                    IDkennung = IDkennung & "#"
                                End Try

                            Next

                        End If

                    Case PTpfdk.PhaseCategories

                        For i = 1 To mycollection.Count

                            cName = splitHryFullnameTo1(CStr(mycollection.Item(i)))
                            'cName = CStr(mycollection.Item(i)).Replace("#", "-")
                            ' der evtl vorhandenen Breadcrumb hat als Trennzeichen das #
                            Try
                                IDkennung = IDkennung & "#" & cName
                            Catch ex As Exception
                                IDkennung = IDkennung & "#"
                            End Try

                        Next

                    Case PTpfdk.Meilenstein

                        For i = 1 To mycollection.Count
                            ' Änderung tk 30.5.17
                            cName = splitHryFullnameTo1(CStr(mycollection.Item(i)))
                            'cName = CStr(mycollection.Item(i)).Replace("#", "-")
                            IDkennung = IDkennung & "#" & cName

                        Next


                    Case PTpfdk.MilestoneCategories
                        For i = 1 To mycollection.Count
                            ' Änderung tk 30.5.17
                            cName = splitHryFullnameTo1(CStr(mycollection.Item(i)))
                            'cName = CStr(mycollection.Item(i)).Replace("#", "-")
                            IDkennung = IDkennung & "#" & cName

                        Next

                    Case PTpfdk.Rollen

                        If mycollection.Count = RoleDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count
                                cName = CStr(mycollection.Item(i))
                                ' bei den cNames ist es jetzt roleUid;teamUid bzw roleUid; von daher einfach umverändert übernehmen 
                                'IDkennung = IDkennung & "#" & RoleDefinitions.getRoledef(cName).UID.ToString
                                IDkennung = IDkennung & "#" & cName
                            Next

                        End If

                    Case PTpfdk.Skill

                        If mycollection.Count = RoleDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count
                                cName = CStr(mycollection.Item(i))
                                ' bei den cNames ist es jetzt roleUid;teamUid bzw roleUid; von daher einfach umverändert übernehmen 
                                'IDkennung = IDkennung & "#" & RoleDefinitions.getRoledef(cName).UID.ToString
                                IDkennung = IDkennung & "#" & cName
                            Next

                        End If

                    Case PTpfdk.Kosten

                        If mycollection.Count = CostDefinitions.Count Then
                            IDkennung = IDkennung & "#Alle"

                        ElseIf CInt(mycollection.Count) = 1 And CStr(mycollection.Item(1)) = "TotalCost" Then
                            IDkennung = IDkennung & "#Alle"

                        Else

                            For i = 1 To mycollection.Count
                                cName = CStr(mycollection.Item(i))
                                If CostDefinitions.containsName(cName) Then
                                    IDkennung = IDkennung & "#" & CostDefinitions.getCostdef(cName).UID.ToString
                                End If

                            Next

                        End If

                    Case PTpfdk.ErgebnisWasserfall

                        If mycollection.Count > 0 Then
                            cName = CStr(mycollection.Item(1))
                            IDkennung = IDkennung & "#" & cName
                        End If

                    Case PTpfdk.Budget

                        If mycollection.Count > 0 Then
                            cName = CStr(mycollection.Item(1))
                            IDkennung = IDkennung & "#" & cName
                        End If

                    Case PTpfdk.Cashflow

                        If mycollection.Count > 0 Then
                            cName = CStr(mycollection.Item(1))
                            IDkennung = IDkennung & "#" & cName
                        End If
                End Select
            Catch ex As Exception

                IDkennung = IDkennung & "#?"
            End Try



        ElseIf typ = "pr" Then

            IDkennung = IDkennung & "#" & CStr(mycollection.Item(1))
            If mycollection.Count >= 2 Then
                IDkennung = IDkennung & "#" & CStr(mycollection.Item(2))
            End If

        End If


        calcChartKennung = IDkennung


    End Function
    ''' <summary>
    ''' errechnet den für Showprojekte und AlleProjekte benötigten Schlüssel
    ''' setzt sich zusammen aus pName und variantName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjektKey(ByVal pName As String, ByVal variantName As String) As String

        Dim trennzeichen As String = "#"

        ' Konsistenzbedingungen gewährleisten
        If IsNothing(pName) Then
            Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
        ElseIf pName.Length < 2 Then
            Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & pName)
        ElseIf IsNothing(variantName) Then
            variantName = ""
        End If

        calcProjektKey = pName & trennzeichen & variantName


    End Function

    ''' <summary>
    ''' berechnet den Key für eine customUserRole , setzt sich zusammen aus userName, Kennung CustomRole und ggf specifics, 
    ''' falls es sich um eine RessourceManager Rolle handelt 
    ''' </summary>
    ''' <param name="userName"></param>
    ''' <param name="customRoleType"></param>
    ''' <param name="specifics"></param>
    ''' <returns></returns>
    Public Function calcCurKey(ByVal userName As String, ByVal customRoleType As ptCustomUserRoles, ByVal specifics As String) As String

        Dim key As String = userName.Trim & CInt(customRoleType).ToString.Trim
        If customRoleType = ptCustomUserRoles.RessourceManager Or
           customRoleType = ptCustomUserRoles.TeamManager Or
           customRoleType = ptCustomUserRoles.InternalViewer Or
           customRoleType = ptCustomUserRoles.ExternalViewer Then

            key = key & specifics

        End If

        calcCurKey = key

    End Function


    ''' <summary>
    ''' eine Zahl wird in einen achstelligen String mit führenden Nullen gewandelt, 
    ''' damit das im Falle customTF als sortkey verwendet werden kann 
    ''' </summary>
    ''' <param name="zeile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcSortKeyCustomTF(ByVal zeile As Integer) As String
        If zeile >= 2 Then
            calcSortKeyCustomTF = zeile.ToString("00000000")
        Else
            zeile = 2
            calcSortKeyCustomTF = zeile.ToString("00000000")
        End If

    End Function

    ''' <summary>
    ''' falls ein TF-key bereits existiert, wird durch Append von "x" ein neuer key erzeugt
    ''' </summary>
    ''' <param name="oldStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcSortKeyCustomTF1(ByVal oldStr As String) As String
        calcSortKeyCustomTF1 = oldStr & "x"
    End Function

    ''' <summary>
    ''' sorgt dafür. daß Projekte immer im gleichen Muster angezeigt werden 
    ''' Erst sortiert nach BU, dann nach ProjektStart-Datum, dann nach Länge  
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcKennziffer(ByVal hproj As clsProjekt) As Double

        Dim wertigkeitBU As Integer = 100000
        Dim wertigkeitDate As Double = 100
        Dim wertigkeitLaenge As Double = 0.1
        Dim zwErg As Double = 0.0

        Dim found As Boolean = False
        Dim i As Integer = 1

        While i <= businessUnitDefinitions.Count And Not found

            If businessUnitDefinitions.ElementAt(i - 1).Value.name = hproj.businessUnit Then
                found = True
            Else
                i = i + 1
            End If

        End While

        zwErg = i * wertigkeitBU

        ' Berücksichtigung ProjektstartDatum 
        zwErg = zwErg + DateDiff(DateInterval.Day, StartofCalendar, hproj.startDate) / 30.4 * wertigkeitDate

        ' Berücksichtigung Länge
        zwErg = zwErg + hproj.dauerInDays / 30.4 * wertigkeitLaenge

        calcKennziffer = zwErg

    End Function

    ''' <summary>
    ''' Funktion, um die Kalenderwoche zu bestimmen 
    ''' </summary>
    ''' <param name="datum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcKW(ByVal datum As Date) As Integer

        Dim kw As Integer

        kw = DatePart(DateInterval.WeekOfYear, datum, FirstDayOfWeek.Monday,
          FirstWeekOfYear.FirstFourDays)

        calcKW = kw

    End Function

    ''' <summary>
    ''' calculates sum of values starting at 0 until and including index  
    ''' </summary>
    ''' <param name="ar"></param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public Function calcPartSum2ix(ByVal ar As Double(), ByVal index As Integer) As Double
        Dim result As Double = 0

        Dim arLength As Integer = ar.Length
        If index >= arLength Or index < 0 Then
            ' do nothing 
        Else
            For ix As Integer = 0 To index
                result = result + ar(ix)
            Next
        End If

        calcPartSum2ix = result
    End Function

    ''' <summary>
    ''' calculates sum of values starting with and including index until end of array 
    ''' </summary>
    ''' <param name="ar"></param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public Function calcPartSum2End(ByVal ar As Double(), ByVal index As Integer) As Double
        Dim result As Double = 0
        Dim arLength As Integer = ar.Length
        If index >= arLength Or index < 0 Then
            ' do nothing 
        Else
            For ix As Integer = index To arLength - 1
                result = result + ar(ix)
            Next
        End If

        calcPartSum2End = result
    End Function

    ''' <summary>
    ''' gibt den Standard-Last Complete Session Scenario Namen des Nutzers zurück 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcLastSessionScenarioName() As String
        Dim tmpResult As String = "_last Session (by " & dbUsername & ")"
        tmpResult = calcPortfolioKey(tmpResult, "")
        calcLastSessionScenarioName = tmpResult
    End Function

    ' tk geändert: es gibt nur noch die lastSession und eine gespeicherte Konstellation
    ' ''' <summary>
    ' ''' gibt den Standard last Editor Szenario Namen zurück 
    ' ''' </summary>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function calcLastEditorScenarioName() As String
    '    Dim tmpResult As String = "_last Scenario-Editor (by " & dbUsername & ")"
    '    calcLastEditorScenarioName = tmpResult
    'End Function

    ''' <summary>
    ''' errechnet den Namen, den das Text Shape eines Projektes hat; Input ist der Projekt-Name
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjectTextShapeName(ByVal pName As String) As String
        Dim tmpResult As String = "DummyName"
        If Not IsNothing(pName) Then
            tmpResult = "t0a1" & pName
        End If
        calcProjectTextShapeName = tmpResult
    End Function

    ''' <summary>
    ''' errechnet den für Showprojekte und AlleProjekte benötigten Schlüssel
    ''' verwendet dazu die in hproj vorhandenen Attribute Name und variantName
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjektKey(ByVal hproj As clsProjekt) As String

        Dim trennzeichen As String = "#"


        ' Konsistenzbedingungen gewährleisten
        If IsNothing(hproj.name) Then
            Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
        ElseIf hproj.name.Length < 2 Then
            Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & hproj.name)
        ElseIf IsNothing(hproj.variantName) Then
            hproj.variantName = ""
        End If

        calcProjektKey = hproj.name & trennzeichen & hproj.variantName

    End Function

    ''' <summary>
    ''' berechnet die ID für die Datenbank, bestehend aus Projektname, Variant-Name und TimeStamp
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="datum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcProjektUidDB(ByVal hproj As clsProjekt, ByVal datum As Date) As String

        Dim trennzeichen As String = "#"

        With hproj

            ' Konsistenzbedingungen gewährleisten
            If IsNothing(.name) Then
                Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
            ElseIf .name.Length < 2 Then
                Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & .name)
            ElseIf IsNothing(.variantName) Then
                .variantName = ""
            End If

            If IsNothing(datum) Then
                datum = Date.Now
            End If

            calcProjektUidDB = .name & trennzeichen & .variantName & trennzeichen & datum.ToString

        End With


    End Function


    Public Function calcProjektKeyDB(ByVal pName As String, ByVal vName As String) As String

        Dim tmpName As String

        ' Konsistenzbedingungen gewährleisten
        If IsNothing(pName) Then
            Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
        ElseIf pName.Length < 2 Then
            Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & pName)
        ElseIf IsNothing(vName) Then
            vName = ""
        End If

        If vName <> "" And vName.Trim.Length > 0 Then
            tmpName = calcProjektKey(pName, vName)
        Else
            tmpName = pName
        End If

        calcProjektKeyDB = tmpName
    End Function

    ''' <summary>
    ''' errechnet den für projectConstellations benötigten Schlüssel
    ''' setzt sich zusammen aus pName und variantName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcPortfolioKey(ByVal pName As String, ByVal variantName As String) As String

        Dim trennzeichen As String = "#"

        ' Konsistenzbedingungen gewährleisten
        If IsNothing(pName) Then
            Throw New ArgumentException("Projekt-Name kann nicht Nothing sein")
        ElseIf pName.Length < 2 Then
            Throw New ArgumentException("Projekt-Name muss mindestens zwei Zeichen lang sein: " & pName)
        ElseIf IsNothing(variantName) Then
            variantName = ""
        End If

        calcPortfolioKey = pName & trennzeichen & variantName

    End Function

    ''' <summary>
    ''' errechnet den für projektConstellations benötigten Schlüssel
    ''' verwendet dazu die in portfolio vorhandenen Attribute constellationName und variantName
    ''' </summary>
    ''' <param name="portfolio"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcPortfolioKey(ByVal portfolio As clsConstellation) As String

        Dim trennzeichen As String = "#"
        With portfolio

            ' Konsistenzbedingungen gewährleisten
            If IsNothing(.constellationName) Then
                Throw New ArgumentException("Portfolio-Name kann nicht Nothing sein")
            ElseIf .constellationName.Length < 2 Then
                Throw New ArgumentException("Portfolio-Name muss mindestens zwei Zeichen lang sein: " & .constellationName)
            ElseIf IsNothing(.variantName) Then
                .variantName = ""
            End If

            calcPortfolioKey = .constellationName & trennzeichen & .variantName

        End With


    End Function

    ''' <summary>
    ''' berechnet die Koordinaten des Abhängigkeit-Konnektors - der Linie
    ''' </summary>
    ''' <param name="pShape"></param>
    ''' <param name="dpShape"></param>
    ''' <param name="X1"></param>
    ''' <param name="Y1"></param>
    ''' <param name="X2"></param>
    ''' <param name="Y2"></param>
    ''' <remarks></remarks>
    Public Sub calculateDepCoord(ByVal pShape As Excel.Shape, ByVal dpShape As Excel.Shape,
                                     ByRef X1 As Single, ByRef Y1 As Single, ByRef X2 As Single, ByRef Y2 As Single)

        With pShape
            X1 = .Left + .Width / 2
            Y1 = .Top + .Height
        End With

        With dpShape
            X2 = .Left + .Width / 2
            Y2 = .Top
        End With

    End Sub



    Public Sub awinCalculateOptimization(ByVal diagrammTyp As String, ByRef myCollection As Collection,
                                  ByRef OptimierungsErgebnis As SortedList(Of String, clsOptimizationObject))
        Dim referenceValue As Double, newReferenceValue As Double, currentValue As Double
        Dim bestValue As Double
        Dim startoffset() As Integer
        Dim versatz As Integer
        Dim ErgebnisListe As New SortedList(Of Double, clsOptimizationObject)
        'Dim ErgebnisListe As New SortedDictionary(Of Double, clsOptimizationObject)
        Dim lokalesOptimum As clsOptimizationObject
        Dim hproj As clsProjekt
        Dim saveOffset As Integer, anzahlVersuche As Integer
        Dim NrLoops As Integer
        Dim NrArgExceptions As Integer
        Dim avgValue As Double

        ReDim startoffset(0)

        If diagrammTyp = DiagrammTypen(0) Then ' Phase 
            Call MsgBox("Phasen Optimierung noch nicht implementiert")
        ElseIf diagrammTyp = DiagrammTypen(1) Or diagrammTyp = DiagrammTypen(2) Then

            With ShowProjekte
                If diagrammTyp = DiagrammTypen(1) Then
                    referenceValue = .getbadCostOfRole(myCollection)
                Else
                    referenceValue = .getDeviationfromAverage(myCollection, avgValue, diagrammTyp)
                End If

                newReferenceValue = -1
                NrLoops = 0
                NrArgExceptions = 0


                While newReferenceValue < referenceValue And NrLoops < 5 * ShowProjekte.Count
                    ' notwendig für den zweiten , ..n. durchlauf 
                    If newReferenceValue >= 0 Then
                        referenceValue = newReferenceValue
                    End If

                    For Each kvp As KeyValuePair(Of String, clsProjekt) In .Liste

                        If relevantForOptimization(kvp.Value) Then

                            bestValue = referenceValue  ' als Startwert, der hoffentlich unterboten wird .... 
                            startoffset(0) = 0

                            For versatz = kvp.Value.earliestStart To kvp.Value.latestStart
                                If versatz <> 0 Then
                                    kvp.Value.StartOffset = versatz
                                    If diagrammTyp = DiagrammTypen(1) Then
                                        currentValue = .getbadCostOfRole(myCollection)
                                    Else
                                        currentValue = .getDeviationfromAverage(myCollection, avgValue, diagrammTyp)
                                    End If

                                    If currentValue < bestValue Then
                                        bestValue = currentValue
                                        startoffset(0) = versatz
                                    End If
                                End If
                            Next versatz

                            ' zurücksetzen des StartOffsets im Projekt, weil hier ja erst verschiedene Konstellationen probiert werden  
                            kvp.Value.StartOffset = 0

                            If startoffset(0) <> 0 Then ' es gab eine Verbesserung 
                                lokalesOptimum = New clsOptimizationObject
                                With lokalesOptimum
                                    .projectName = kvp.Key
                                    '.bestValue = bestValue
                                    .offset = startoffset
                                End With

                                Try
                                    ErgebnisListe.Add(bestValue, lokalesOptimum)
                                Catch ex As ArgumentException
                                    NrArgExceptions = NrArgExceptions + 1
                                    bestValue = bestValue + NrArgExceptions * 0.00000017
                                    Try
                                        ErgebnisListe.Add(bestValue, lokalesOptimum)
                                    Catch ex1 As ArgumentException
                                        NrArgExceptions = NrArgExceptions + 1
                                        bestValue = bestValue + NrArgExceptions * 0.00000017
                                        ErgebnisListe.Add(bestValue, lokalesOptimum)
                                    End Try
                                End Try
                            End If
                        End If

                    Next kvp
                    '
                    ' jetzt muss die Ergebnis Liste abgearbeitet werden ... 
                    '
                    anzahlVersuche = 0
                    newReferenceValue = referenceValue

                    For Each ergebnis As KeyValuePair(Of Double, clsOptimizationObject) In ErgebnisListe

                        hproj = ShowProjekte.getProject(ergebnis.Value.projectName)
                        saveOffset = hproj.StartOffset
                        hproj.StartOffset = ergebnis.Value.offset(0)

                        If diagrammTyp = DiagrammTypen(1) Then
                            currentValue = .getbadCostOfRole(myCollection)
                        Else
                            currentValue = .getDeviationfromAverage(myCollection, avgValue, diagrammTyp)
                        End If

                        If currentValue < newReferenceValue Then
                            newReferenceValue = currentValue
                            anzahlVersuche = 0
                            ' hier müssen best, second, third gesetzt werden
                        Else
                            hproj.StartOffset = saveOffset
                            anzahlVersuche = anzahlVersuche + 1
                            If anzahlVersuche > 5 Or ergebnis.Value.offset(0) = 0 Then
                                ' wenn startoffset = 0 , dann konnten keine Verbesserungen mehr erzielt werden , also Abbruch ...
                                Exit For
                            End If
                        End If
                        NrLoops = NrLoops + 1

                    Next

                    ErgebnisListe.Clear()

                End While
                ' hier wird Gold gesetzt, das heißt alle Offsets gemerkt, die für die Optmierung notwendig sind 
                ' anschließend werden alle startoffsets wieder auf 0 (=Ausgangswert) gesetzt 
                Dim tmpOffset() As Integer
                ReDim tmpOffset(0)
                OptimierungsErgebnis.Clear()
                For Each kvp As KeyValuePair(Of String, clsProjekt) In .Liste
                    If kvp.Value.StartOffset <> 0 Then
                        lokalesOptimum = New clsOptimizationObject
                        With lokalesOptimum
                            .projectName = kvp.Value.name
                            '.bestValue = bestValue
                            tmpOffset(0) = kvp.Value.StartOffset
                            .offset = tmpOffset
                        End With
                        OptimierungsErgebnis.Add(kvp.Value.name, lokalesOptimum)
                        'kvp.Value.StartOffset = 0
                    End If
                Next kvp

            End With

            'Case DiagrammTypen(2) ' Cost
            '    Call MsgBox("Kosten-Diagramm Optimierung noch nicht implementiert")
        Else
            Call MsgBox("Sonstige Optimierung noch nicht implementiert")
        End If

    End Sub

    Public Sub awinCalcOptimizationVarianten(ByVal diagrammTyp As String, ByRef myCollection As Collection,
                                             ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        Dim anzahlVarianten As Integer
        Dim maxValue() As Integer
        Dim indexValue() As Integer
        Dim anzProjMitVar As Integer
        Dim PPointer As Integer
        Dim anzSchleifen As Integer = 0
        Dim firstValue As Double = 100000000000.0
        Dim secondValue As Double = 100000000000.0
        Dim thirdValue As Double = 100000000000.0
        Dim atleastOne As Boolean = False
        Dim anzKombinationen As Integer = 1
        Dim anzOptimierungen As Integer = 0

        Dim moreThanOne As New Collection
        Dim justOne As New Collection

        ' bestimme die Collection mit Projekten mit mehr als einer Variante
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            anzahlVarianten = AlleProjekte.getVariantNames(kvp.Key, True).Count
            If anzahlVarianten = 1 Then
                justOne.Add(kvp.Key, kvp.Key)
            ElseIf anzahlVarianten > 1 Then
                moreThanOne.Add(kvp.Key, kvp.Key)
            End If
        Next


        If moreThanOne.Count = 0 Then
            e.Result = "es gibt keine Varianten .. demnach gibt es auch nichts zu optimieren !"
            worker.ReportProgress(0, e)
            Exit Sub
        Else
            ' speichern der letzten Konstellation
            Call storeSessionConstellation(autoSzenarioNamen(0))
        End If


        ' nimmt die Anzahl der Varianten auf
        anzProjMitVar = moreThanOne.Count
        ReDim maxValue(anzProjMitVar - 1)
        ReDim indexValue(anzProjMitVar - 1)

        ' jetzt wird bestimmt: 
        ' wievele Varianten hat das i.-te Element in morethanOne
        ' an welcher Stelle steht der Varianten-Zeiger Zeiger für das die bestimmt werden 
        Dim i As Integer = 0

        anzKombinationen = 1
        For Each pName As String In moreThanOne
            maxValue(i) = AlleProjekte.getVariantZahl(pName)
            ' in maxvalue steht 0, wenn es nur die Basis Variante gibt ..
            If maxValue(i) >= 0 Then
                anzKombinationen = anzKombinationen * (maxValue(i) + 1)
            End If

            indexValue(i) = 0
            i = i + 1
        Next

        PPointer = 0

        ' Start der Rekursion - und die Ausgangs-Konstellation als Vorgabe, als aktuelle "1. Varianten Optimum" behalten 
        firstValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
        Call storeSessionConstellation(autoSzenarioNamen(1))

        ' die anderen Szenarien sollen jetzt gelöscht werden 
        If projectConstellations.Contains(autoSzenarioNamen(2)) Then
            projectConstellations.Remove(autoSzenarioNamen(2))
        End If

        If projectConstellations.Contains(autoSzenarioNamen(3)) Then
            projectConstellations.Remove(autoSzenarioNamen(3))
        End If



        Call IterateOptimization(PPointer, anzProjMitVar, maxValue, indexValue,
                                diagrammTyp, myCollection, anzKombinationen, anzSchleifen, anzOptimierungen,
                                justOne, moreThanOne,
                                firstValue, secondValue, thirdValue,
                                worker, e)



        If anzOptimierungen > 0 Then
            ' wieder den alten Zustand herstellen 
            Call loadSessionConstellation(calcPortfolioKey(autoSzenarioNamen(0), ""), False, True)
        Else
            ' es hat sich eh nichts geändert ... 
            'Call loadSessionConstellation(autoSzenarioNamen(0), False, False)
            e.Result = "in " & anzSchleifen.ToString & " Kombinationen" & vbLf & "konnte keine Verbesserung gefunden werden"
            worker.ReportProgress(0, e)

        End If


        ' erstelle alle Kombinationen der Varianten in der Variablen Current 


    End Sub




    ''' <summary>
    ''' rekursive Funktion, die die Kombinatorik der Varianten ermittelt 
    ''' </summary>
    ''' <param name="PPointer"></param>
    ''' <param name="anzProjMitVar"></param>
    ''' <param name="maxvalue"></param>
    ''' <param name="indexvalue"></param>
    ''' <param name="anzSchleifen"></param>
    ''' <remarks></remarks>
    Private Sub IterateOptimization(ByVal PPointer As Integer, ByVal anzProjMitVar As Integer,
                                           ByVal maxvalue() As Integer, ByVal indexvalue() As Integer,
                                           ByRef diagrammTyp As String, ByRef myCollection As Collection,
                                           ByVal anzKombinationen As Integer, ByRef anzSchleifen As Integer, ByRef anzOptimierungen As Integer,
                                           ByRef justOne As Collection, ByRef moreThanOne As Collection,
                                           ByRef firstValue As Double, ByRef secondValue As Double, ByRef thirdValue As Double,
                                           ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        'Dim currentSzenario As New clsProjekte
        Dim currentValue As Double
        Dim tmpConstellation As clsConstellation


        Dim hproj As clsProjekt

        If worker.CancellationPending Then
            e.Cancel = True
            e.Result = "Optimierung  abgebrochen ..."
            Exit Sub
        End If


        If PPointer = anzProjMitVar - 1 Then

            indexvalue(anzProjMitVar - 1) = 0

            While indexvalue(anzProjMitVar - 1) <= maxvalue(anzProjMitVar - 1)


                ' jetzt die Aktion ausführen 

                Dim txtMSG As String = ""
                'currentSzenario = New clsProjekte
                For i = 1 To anzProjMitVar

                    If i = 1 Then
                        txtMSG = indexvalue(i - 1).ToString & ", "
                    ElseIf i = anzProjMitVar Then
                        txtMSG = txtMSG & indexvalue(i - 1).ToString
                    Else
                        txtMSG = txtMSG & indexvalue(i - 1).ToString & ", "
                    End If

                    hproj = AlleProjekte.getProject(CStr(moreThanOne.Item(i)), indexvalue(i - 1))
                    'currentSzenario.Add(hproj)

                    ShowProjekte.AddAnyway(hproj)
                    'Call replaceProjectVariant(hproj.name, hproj.variantName, False, False, 0)

                Next


                ' jetzt muss der Wert für current bestimmt werden 
                currentValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
                anzSchleifen = anzSchleifen + 1

                If currentValue < firstValue Then
                    thirdValue = secondValue
                    secondValue = firstValue
                    firstValue = currentValue

                    anzOptimierungen = anzOptimierungen + 1


                    If projectConstellations.Contains(autoSzenarioNamen(2)) Then
                        tmpConstellation = projectConstellations.getConstellation(autoSzenarioNamen(2))

                        If projectConstellations.Contains(autoSzenarioNamen(3)) Then
                            projectConstellations.Remove(autoSzenarioNamen(3))
                        End If

                        tmpConstellation.constellationName = autoSzenarioNamen(3)
                        projectConstellations.Add(tmpConstellation)

                        If projectConstellations.Contains(autoSzenarioNamen(1)) Then
                            tmpConstellation = projectConstellations.getConstellation(autoSzenarioNamen(1))

                            If projectConstellations.Contains(autoSzenarioNamen(2)) Then
                                projectConstellations.Remove(autoSzenarioNamen(2))
                            End If

                            tmpConstellation.constellationName = autoSzenarioNamen(2)
                            projectConstellations.Add(tmpConstellation)
                        End If


                    End If

                    Call storeSessionConstellation(autoSzenarioNamen(1))
                    'Call awinNeuZeichnenDiagramme(2)

                ElseIf currentValue < secondValue Then

                    anzOptimierungen = anzOptimierungen + 1

                    thirdValue = secondValue
                    secondValue = currentValue

                    If projectConstellations.Contains(autoSzenarioNamen(2)) Then
                        tmpConstellation = projectConstellations.getConstellation(autoSzenarioNamen(2))

                        If projectConstellations.Contains(autoSzenarioNamen(3)) Then
                            projectConstellations.Remove(autoSzenarioNamen(3))
                        End If

                        tmpConstellation.constellationName = autoSzenarioNamen(3)
                        projectConstellations.Add(tmpConstellation)

                    End If

                    Call storeSessionConstellation(autoSzenarioNamen(2))

                ElseIf currentValue < thirdValue Then

                    anzOptimierungen = anzOptimierungen + 1

                    thirdValue = currentValue
                    Call storeSessionConstellation(autoSzenarioNamen(3))
                End If

                e.Result = anzSchleifen.ToString & " / " & anzKombinationen.ToString & " Berechnungen; " &
                            anzOptimierungen.ToString & " Optimierung(en"
                worker.ReportProgress(0, e)
                indexvalue(PPointer) = indexvalue(PPointer) + 1


            End While

            indexvalue(anzProjMitVar - 1) = 0


        Else

            For i = 0 To maxvalue(PPointer)
                indexvalue(PPointer) = i
                Call IterateOptimization(PPointer + 1, anzProjMitVar, maxvalue, indexvalue,
                                        diagrammTyp, myCollection, anzKombinationen, anzSchleifen, anzOptimierungen,
                                        justOne, moreThanOne, firstValue, secondValue, thirdValue,
                                        worker, e)

                If worker.CancellationPending Then
                    e.Cancel = True
                    e.Result = "Optimierung  abgebrochen ..."
                    Exit For
                End If

            Next


        End If




    End Sub

    ''' <summary>
    ''' bereichnet auf Basis der Freiheitsgrade der Projekte die beste Konstellation
    ''' </summary>
    ''' <param name="diagrammTyp"></param>
    ''' <param name="myCollection"></param>
    ''' <param name="OptimierungsErgebnis"></param>
    ''' <remarks></remarks>
    Public Sub awinCalcOptimizationFreiheitsgrade(ByVal diagrammTyp As String, ByRef myCollection As Collection,
                                       ByRef OptimierungsErgebnis As SortedList(Of String, clsOptimizationObject))
        Dim currentValue As Double
        Dim bestValue As Double
        Dim startoffset As Integer
        Dim versatz As Integer
        Dim lokalesOptimum As New clsOptimizationObject
        Dim hproj As clsProjekt
        Dim NrArgExceptions As Integer
        Dim toDoListe As New Collection
        Dim NrLoops As Integer


        If myCollection.Count >= 1 Then


            If diagrammTyp = DiagrammTypen(0) Or diagrammTyp = DiagrammTypen(1) Or diagrammTyp = DiagrammTypen(2) Or diagrammTyp = DiagrammTypen(4) Then

                ' to do Liste aufbauen
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If relevantForOptimization(kvp.Value) Then
                        toDoListe.Add(kvp.Key, kvp.Key)
                    End If
                Next kvp

                bestValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
                lokalesOptimum.bestValue = bestValue
                lokalesOptimum.projectName = " "
                OptimierungsErgebnis.Clear()

                NrLoops = 0
                NrArgExceptions = 0


                'While newReferenceValue < referenceValue And toDoListe.Count > 0
                Dim Abbruch As Boolean = False
                While toDoListe.Count > 0 And Not Abbruch

                    Dim i As Integer
                    Dim curProj As clsProjekt

                    For i = 1 To toDoListe.Count
                        curProj = ShowProjekte.getProject(CStr(toDoListe.Item(i)))

                        startoffset = 0

                        ' hier wird der beste Wert für das einzelne Projekt gesucht ....  

                        For versatz = curProj.earliestStart To curProj.latestStart
                            If versatz <> 0 Then
                                curProj.StartOffset = versatz
                                currentValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)

                                If currentValue < bestValue Then
                                    bestValue = currentValue
                                    startoffset = versatz
                                End If
                            End If
                        Next versatz

                        ' zurücksetzen des StartOffsets im Projekt, weil hier ja erst verschiedene Konstellationen probiert werden  
                        curProj.StartOffset = 0

                        Dim tmpOffset() As Integer
                        ReDim tmpOffset(0)

                        If startoffset <> 0 Then ' es gab eine Verbesserung 
                            'lokalesOptimum = New clsOptimizationObject
                            With lokalesOptimum
                                If bestValue < .bestValue Then
                                    .projectName = curProj.name
                                    .bestValue = bestValue
                                    tmpOffset(0) = startoffset
                                    .offset = tmpOffset
                                    ' Call awinVisualizeProject
                                End If
                            End With

                        End If

                    Next i
                    '
                    ' jetzt muss das Ergebnis abgearbeitet werden ... 
                    '
                    If lokalesOptimum.projectName <> " " Then

                        hproj = ShowProjekte.getProject(lokalesOptimum.projectName)
                        hproj.StartOffset = lokalesOptimum.offset(0)
                        OptimierungsErgebnis.Add(lokalesOptimum.projectName, lokalesOptimum)
                        toDoListe.Remove(lokalesOptimum.projectName)
                        Call visualisiereTeilErgebnis(lokalesOptimum.projectName)
                    Else
                        Abbruch = True
                    End If

                    lokalesOptimum.projectName = " "
                    NrLoops = NrLoops + 1

                End While

            Else
                Call MsgBox("Optimierung noch nicht implementiert")
            End If
        Else
            Call MsgBox("Optimierung nicht implementiert")
        End If


    End Sub


    ''' <summary>
    ''' bereichnet auf Basis der Freiheitsgrade der Plan-Elemente Elemente  die beste Konstellation
    ''' </summary>
    ''' <param name="diagrammTyp"></param>
    ''' <param name="myCollection"></param>
    ''' <param name="OptimierungsErgebnis"></param>
    ''' <remarks></remarks>
    Public Sub awinCalcOptimizationElemFreiheitsgrade(ByVal diagrammTyp As String, ByVal myCollection As Collection,
                                       ByRef OptimierungsErgebnis As SortedList(Of String, clsOptimizationObject),
                                       ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)
        Dim currentValue As Double
        Dim fullname As String = ""
        Dim elemName As String = ""
        Dim elemID As String
        Dim breadcrumb As String = ""
        Dim bestValue As Double
        Dim versatz As Integer
        Dim lokalesOptimum As New clsOptimizationObject
        Dim deltaValues() As Integer
        Dim hproj As clsProjekt
        Dim NrArgExceptions As Integer
        Dim toDoListe As New Collection
        Dim NrLoops As Integer
        Dim cphase As clsPhase
        Dim zaehler As Integer = 1
        Dim anzImprovements As Integer = 0
        Dim backgroundMsg As String = ""
        Dim i As Integer
        Dim notAdded As Boolean = True
        Dim phaseIndices() As Integer


        If myCollection.Count = 1 Then


            If diagrammTyp = DiagrammTypen(0) And myCollection.Count = 1 Then

                ' to do Liste aufbauen
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    ' tk 18.11.15 checken ob eines der angegebenen Elemente vorkommt 
                    notAdded = True
                    fullname = CStr(myCollection.Item(1))

                    elemName = ""
                    breadcrumb = ""
                    Dim pvName As String = ""
                    Dim type As Integer = -1
                    Dim weiter As Boolean = True
                    Call splitHryFullnameTo2(fullname, elemName, breadcrumb, type, pvName)

                    If type = PTItemType.projekt Then

                        If pvName <> calcProjektKey(kvp.Value) Then
                            weiter = False
                        End If

                    ElseIf type = PTItemType.vorlage Then

                        ' anyway ok

                    End If

                    If weiter Then
                        phaseIndices = kvp.Value.hierarchy.getPhaseIndices(elemName, breadcrumb)
                        If phaseIndices(0) > 0 Then

                            i = 1
                            While i <= phaseIndices.Length And notAdded
                                elemID = kvp.Value.getPhase(phaseIndices(i - 1)).nameID
                                If relevantForOptimization(kvp.Value, elemID, False) Then
                                    If Not toDoListe.Contains(kvp.Key) Then
                                        toDoListe.Add(kvp.Key, kvp.Key)
                                    End If
                                    notAdded = False
                                Else
                                    i = i + 1
                                End If
                            End While

                        End If
                    End If

                Next kvp

                bestValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)
                lokalesOptimum.bestValue = bestValue
                lokalesOptimum.projectName = " "
                OptimierungsErgebnis.Clear()

                NrLoops = 0
                NrArgExceptions = 0


                Dim Abbruch As Boolean = False
                While toDoListe.Count > 0 And Not Abbruch

                    backgroundMsg = "Iteration " & zaehler.ToString("###0") &
                                    "; gefundene Verbesserungen: " & anzImprovements.ToString("###0")

                    e.Result = backgroundMsg
                    worker.ReportProgress(0, e)

                    Dim curProj As clsProjekt
                    fullname = CStr(myCollection.Item(1))

                    elemName = ""
                    breadcrumb = ""
                    Dim pvName As String = ""
                    Dim type As Integer = -1
                    Call splitHryFullnameTo2(fullname, elemName, breadcrumb, type, pvName)

                    For i = 1 To toDoListe.Count

                        curProj = ShowProjekte.getProject(CStr(toDoListe.Item(i)))
                        Dim weiter As Boolean = True

                        If type = PTItemType.projekt Then

                            If pvName <> calcProjektKey(curProj) Then
                                weiter = False
                            End If

                        ElseIf type = PTItemType.vorlage Then

                            ' anyway weiter

                        End If

                        If weiter Then
                            phaseIndices = curProj.hierarchy.getPhaseIndices(elemName, breadcrumb)

                            ' in den deltaValues sind jetzt die Werte drin, die sich für die Phasen-Verschiebungen ergeben 

                            ReDim deltaValues(phaseIndices.Length - 1)

                            Dim optimizationFound As Boolean = False

                            If phaseIndices(0) > 0 Then
                                ' nur dann wurde die Phase wenigstens einmal gefunden ...
                                For ik As Integer = 1 To phaseIndices.Length

                                    ' jetzt die Phase holen

                                    cphase = curProj.getPhase(phaseIndices(ik - 1))
                                    Dim phaseNameID As String = cphase.nameID

                                    ' hier wird der beste Wert für das einzelne Projekt gesucht ....  

                                    For versatz = curProj.earliestStart To curProj.latestStart
                                        If versatz <> 0 Then
                                            ' jetzt die Phase entsprechend verschieben ...
                                            cphase.changeStartandDauer(cphase.startOffsetinDays + versatz, cphase.dauerInDays)

                                            currentValue = berechneOptimierungsWert(ShowProjekte, diagrammTyp, myCollection)

                                            If currentValue < bestValue Then
                                                bestValue = currentValue
                                                deltaValues(ik - 1) = versatz
                                                optimizationFound = True
                                                anzImprovements = anzImprovements + 1
                                            End If
                                        End If

                                    Next versatz



                                Next

                                ' zurücksetzen des Offsets in den einzelnen Phasen wieder auf ihre alte Position zurückgesetzt werden 

                                For ik As Integer = 1 To phaseIndices.Length
                                    cphase = curProj.getPhase(phaseIndices(ik - 1))
                                    Dim phaseNameID As String = cphase.nameID

                                    If deltaValues(ik - 1) <> 0 Then
                                        With cphase
                                            .changeStartandDauer(.startOffsetinDays - deltaValues(ik - 1), .dauerInDays)
                                        End With
                                    End If

                                Next


                                If optimizationFound Then ' es gab eine Verbesserung 

                                    With lokalesOptimum
                                        If bestValue < .bestValue Then
                                            .projectName = curProj.name
                                            .bestValue = bestValue
                                            .offset = deltaValues
                                            ' Call awinVisualizeProject
                                        End If
                                    End With

                                End If

                            End If

                        End If

                    Next i


                    '
                    ' jetzt muss das Ergebnis abgearbeitet werden ... 
                    '
                    If lokalesOptimum.projectName <> " " Then

                        hproj = ShowProjekte.getProject(lokalesOptimum.projectName)
                        ' jetzt müssen die Phasen wieder auf Ihre optimierte Position gebracht werden

                        phaseIndices = hproj.hierarchy.getPhaseIndices(elemName, breadcrumb)
                        If phaseIndices(0) > 0 Then

                            For ik As Integer = 1 To phaseIndices.Length
                                cphase = hproj.getPhase(phaseIndices(ik - 1))
                                Dim phaseNameID As String = cphase.nameID
                                If lokalesOptimum.offset(ik - 1) <> 0 Then
                                    cphase.changeStartandDauer(cphase.startOffsetinDays + lokalesOptimum.offset(ik - 1),
                                                                cphase.dauerInDays)
                                End If

                            Next

                            OptimierungsErgebnis.Add(lokalesOptimum.projectName, lokalesOptimum)
                            Call visualisiereTeilErgebnis(lokalesOptimum.projectName)

                        End If


                        toDoListe.Remove(lokalesOptimum.projectName)

                    Else
                        Abbruch = True
                    End If

                    lokalesOptimum.projectName = " "
                    NrLoops = NrLoops + 1

                End While

            Else
                Call MsgBox("Optimierung noch nicht implementiert")
            End If
        Else
            Call MsgBox("Optimierung für mehr als 1 Namen noch nicht implementiert")
        End If


    End Sub

End Module
