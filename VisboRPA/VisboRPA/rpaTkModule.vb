Imports xlns = Microsoft.Office.Interop.Excel

Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports DBAccLayer
Imports WebServerAcc
Module rpaTkModule

    ''' <summary>
    ''' stores a Portfolio 
    ''' </summary>
    ''' <param name="projectList"></param>
    ''' <param name="portfolioName"></param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    Private Function storeConstellationFromProjectList(ByVal projectList As clsProjekteAlle,
                                                    ByVal portfolioName As String, ByVal variantName As String) As Boolean

        Dim result = True

        Try

            AlleProjekte.Clear()
            ' now make sure all projects are in AlleProjekte
            For Each ppair As KeyValuePair(Of String, clsProjekt) In projectList.liste
                If Not AlleProjekte.Containskey(ppair.Key) Then
                    AlleProjekte.Add(ppair.Value)
                End If
            Next

            ShowProjekte.Clear()
            For Each ppair As KeyValuePair(Of String, clsProjekt) In projectList.liste

                If Not ShowProjekte.contains(ppair.Value.name) Then
                    ShowProjekte.Add(ppair.Value)
                End If

            Next


            ' currentSessionConstellation is build by alle the Showprojekte.add and AlleProjekte.add Commands ...
            ' create form that a portfolio, only containing the show-Elements 
            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=variantName)

            ' now store the Portfolio , with name portfolioName
            Dim errMsg As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            ' then empty ShowProjekte again 
            ShowProjekte.Clear()

        Catch ex As Exception
            result = False
            Call logger(ptErrLevel.logError, "failure in store Portfolio: " & portfolioName & vbLf & ex.Message, PTRpa.visboProjectList.ToString)
        End Try

        storeConstellationFromProjectList = result

    End Function


    Public Function checkFindBestStarts(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName1 As String = "Find Best Start"
        Dim blattName1A As String = "Find Best Start PM"
        Dim blattName2 As String = "Parameters"


        Try

            Dim currentWS As xlns.Worksheet = Nothing

            Try
                currentWS = CType(currentWB.Worksheets.Item(blattName1), xlns.Worksheet)
            Catch ex As Exception
                currentWS = Nothing
            End Try

            If IsNothing(currentWS) Then
                Try
                    currentWS = CType(currentWB.Worksheets.Item(blattName1A), xlns.Worksheet)
                Catch ex As Exception
                    currentWS = Nothing
                End Try
            End If

            If Not IsNothing(currentWS) Then

                Dim paramWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName2), xlns.Worksheet)

                If IsNothing(currentWS) Or IsNothing(paramWS) Then
                    result = PTRpa.visboUnknown
                Else
                    Dim ersteZeile As xlns.Range = CType(currentWS.Rows.Item(1), xlns.Range)
                    Try

                        verifiedStructure = ersteZeile.Cells(1, 1).value.trim = "Name" And
                            CStr(ersteZeile.Cells(1, 2).value).Trim = "Variant"

                    Catch ex As Exception
                        verifiedStructure = False
                    End Try


                    If verifiedStructure Then

                        If currentWS.Name = blattName1 Then
                            result = PTRpa.visboFindProjectStart
                        Else
                            result = PTRpa.visboFindProjectStartPM
                        End If


                        ' Aktiviere das Worksheet 
                        If CType(currentWB.ActiveSheet, xlns.Worksheet).Name <> currentWS.Name Then
                            currentWS.Activate()
                        End If

                        Dim mymessages As New Collection
                        Dim infomsg As String = "File to find best start dates Phases/Milestones detected: " & currentWB.Name
                        If currentWS.Name = blattName1 Then
                            infomsg = "File to find best start dates Roles/Skills detected: " & currentWB.Name
                        End If
                        Call logger(ptErrLevel.logInfo, infomsg, mymessages)
                        'Console.WriteLine(infomsg)
                    Else
                        result = PTRpa.visboUnknown
                    End If


                End If
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkFindBestStarts = result
    End Function

    Public Function checkAutoAllocate(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim blattName0 As String = "VISBO Auto-Allocate"
        Dim blattName1 As String = "Parameters"

        Dim blattExist(1) As Boolean
        blattExist(0) = False
        blattExist(1) = False


        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name = blattName0) Then
                    blattExist(0) = True
                End If

                If (ws.Name = blattName1) Then
                    blattExist(1) = True
                End If

            Next

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If blattExist(0) And blattExist(1) Then
            result = PTRpa.visboSuggestResourceAllocation
        End If

        checkAutoAllocate = result
    End Function

    Public Function checkAutoAdjustPortfolio(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim blattName0 As String = "Adjustments Ranking List"
        Dim blattName1 As String = "Parameters"


        Dim blattExist(1) As Boolean
        blattExist(0) = False
        blattExist(1) = False

        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name = blattName0) Then
                    blattExist(0) = True
                End If

                If (ws.Name = blattName1) Then
                    blattExist(1) = True
                End If

            Next

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If blattExist(0) And blattExist(1) Then
            result = PTRpa.visboAutoAdjust
        End If

        checkAutoAdjustPortfolio = result
    End Function


    ''' <summary>
    ''' checks whether or not the file is a findFeasiblePortfolio file
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Public Function checkfeasiblePortfolio(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim blattName1 As String = "VISBO"
        Dim blattName2 As String = "Parameters"

        Dim hasVISBO As Boolean = False
        Dim hasParameters As Boolean = False

        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name = blattName1) Then

                    Dim ersteZeile As xlns.Range = CType(ws.Rows.Item(1), xlns.Range)
                    Try
                        hasVISBO = ersteZeile.Cells(1, 1).value.trim = "Name" And
                        CStr(ersteZeile.Cells(1, 2).value).Trim = "Variant"
                    Catch ex As Exception

                    End Try
                End If

                If (ws.Name = blattName2) Then
                    hasParameters = True
                End If

            Next

            If (hasVISBO And hasParameters) Then
                result = PTRpa.visboFindfeasiblePortfolio

                Dim mymessages As New Collection
                Dim infomsg As String = "File to define feasible portfolio detected: " & currentWB.Name
                Call logger(ptErrLevel.logInfo, infomsg, mymessages)
                'Console.WriteLine(infomsg)
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkfeasiblePortfolio = result

    End Function


    ''' <summary>
    ''' creates hedged variants for all the sales Pipeline projects
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Public Function checkCreateHedgedVariants(ByVal currentWB As xlns.Workbook) As PTRpa

        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattPartName As String = "Pipeline"

        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name.Contains(blattPartName)) Then

                    Dim ersteZeile As xlns.Range = CType(ws.Rows.Item(2), xlns.Range)
                    Try
                        If Not IsNothing(ersteZeile.Cells(1, 3).value) Then
                            If IsNumeric(ersteZeile.Cells(1, 3).value) Then
                                Dim tstValue As Double = CDbl(ersteZeile.Cells(1, 3).value)
                                Dim tstDate As Date = CDate(ersteZeile.Cells(1, 4).value)
                                If tstValue > 0 And tstValue <= 1.0 And tstDate <> Date.MinValue Then
                                    verifiedStructure = True
                                    Exit For
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        verifiedStructure = False
                    End Try
                End If


            Next

            If verifiedStructure Then

                result = PTRpa.visboCreateHedgedVariant
                Dim mymessages As New Collection
                Dim infomsg As String = "File to create hedged variants detected: " & currentWB.Name
                Call logger(ptErrLevel.logInfo, infomsg, mymessages)
                'Console.WriteLine(infomsg)
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkCreateHedgedVariants = result

    End Function


    ''' <summary>
    ''' returns empty string for all roles / skills 
    ''' exludedNAmes = true: read line 3, else line 4
    ''' isRoleSkills = true : is is about roles and skill; else: it is about phases and milestones 
    '''  
    ''' </summary>
    ''' <param name="excludedNames"></param>
    ''' <returns></returns>
    Public Function getConsiderationList(ByVal excludedNames As Boolean, ByVal Optional isRoleSkills As Boolean = True) As Collection

        Dim result As New Collection
        Dim blattName As String = "Parameters"

        Dim zeile As Integer = 3
        Dim spalte As Integer = 2
        If excludedNames Then
            zeile = 3
        Else
            zeile = 4
        End If

        Try
            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)
            Dim lastColumn As Integer = CType(currentWS.Cells(zeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlToLeft).Column

            For columnIndex As Integer = 2 To lastColumn
                If Not IsNothing(CType(currentWS.Cells(zeile, columnIndex), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                    Dim myName As String = CStr(CType(currentWS.Cells(zeile, columnIndex), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                    If isRoleSkills Then

                        Dim myNameID As String = ""
                        If RoleDefinitions.containsName(myName) Then
                            Dim myRoleSkill As clsRollenDefinition = RoleDefinitions.getRoledef(myName)
                            If Not IsNothing(myRoleSkill) And Not result.Contains(myName) Then

                                If myRoleSkill.isSkill Then
                                    Dim containingRoleID As Integer = RoleDefinitions.getContainingRoleOfSkillMembers(myRoleSkill.UID).UID
                                    myNameID = RoleDefinitions.bestimmeRoleNameID(containingRoleID, myRoleSkill.UID)
                                Else
                                    Dim skillID As Integer = -1
                                    myNameID = RoleDefinitions.bestimmeRoleNameID(myRoleSkill.UID, skillID)
                                End If

                                result.Add(myNameID, myNameID)

                            End If
                        End If
                    Else
                        If myName <> "" And Not result.Contains(myName) Then
                            result.Add(myName, myName)
                        End If
                    End If

                End If
            Next


        Catch ex As Exception

        End Try

        getConsiderationList = result
    End Function

    Public Function getJobParameters(ByVal blattName As String, ByVal myKennung As PTRpa) As clsJobParameters
        Dim result As New clsJobParameters

        Try

            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                With currentWS

                    Select Case myKennung
                        Case PTRpa.visboFindfeasiblePortfolio

                            result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)


                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(7, 2).value)
                            End If


                        Case PTRpa.visboFindProjectStart

                            result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)

                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            Else
                                result.portfolioName = ""
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            Else
                                result.portfolioVariantName = "new projects"
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            Else
                                result.projectVariantName = "fbs"
                            End If

                            If Not IsNothing(.Cells(8, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(8, 2).value)
                            Else
                                result.defaultLatestEnd = DateSerial(Date.Now.Year + 1, 12, 31)
                            End If

                            If Not IsNothing(.Cells(9, 2).value) Then
                                result.defaultDeltaInDays = CInt(.Cells(9, 2).value)
                            End If

                        Case PTRpa.visboFindProjectStartPM

                            result.limitPhases = CDbl(.Cells(1, 2).value)
                            result.limitMilestones = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.phases = getConsiderationList(True, False)
                            ' zeile 4
                            result.milestones = getConsiderationList(False, False)

                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            Else
                                result.portfolioName = ""
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            Else
                                result.portfolioVariantName = "new projects"
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            Else
                                result.projectVariantName = "fbs"
                            End If

                            If Not IsNothing(.Cells(8, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(8, 2).value)
                            Else
                                result.defaultLatestEnd = DateSerial(Date.Now.Year + 1, 12, 31)
                            End If

                            If Not IsNothing(.Cells(9, 2).value) Then
                                result.defaultDeltaInDays = CInt(.Cells(9, 2).value)
                            End If

                        Case PTRpa.visboSuggestResourceAllocation
                            ' same as feasible Portfolio , except Line 7: project variantName
                            result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)


                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            End If


                        Case PTRpa.visboCreateHedgedVariant

                            ' only Zeile = 5 Portfolio Name is of relevance
                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            End If

                        Case PTRpa.visboAutoAdjust

                            If Not IsNothing((.Cells(1, 2).value)) Then
                                result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            Else
                                ' there is a default Setting in new Method 
                            End If

                            If Not IsNothing((.Cells(2, 2).value)) Then
                                result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)
                            Else
                                ' there is a defualt Setting in new MEthod 
                            End If

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)

                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            Else
                                result.portfolioName = ""
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            Else
                                result.portfolioVariantName = ""
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            Else
                                result.projectVariantName = "auto"
                            End If

                            If Not IsNothing(.Cells(8, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(8, 2).value)
                            Else
                                result.defaultLatestEnd = DateSerial(Date.Now.Year + 1, 12, 31)
                            End If

                            If Not IsNothing(.Cells(9, 2).value) Then
                                result.defaultDeltaInDays = CInt(.Cells(9, 2).value)
                            End If


                    End Select


                End With
            Else
                Call logger(ptErrLevel.logError, "GetJobParameters: missing sheet in File ", currentWB.Name)
            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logsevereError, "GetJobParameters: Severe Error ", ex.Message)
        End Try

        getJobParameters = result

    End Function

    Public Function getPortfolioNames() As String()

        Dim result As String()
        ReDim result(2)
        result(0) = "NoName"
        result(1) = ""
        result(2) = ""

        Dim blattName As String = "Parameters"

        Try

            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                With currentWS
                    If Not IsNothing(.Cells(5, 2).value) Then
                        result(0) = CStr(.Cells(5, 2).value).Trim
                    End If
                    If Not IsNothing(.Cells(6, 2).value) Then
                        result(1) = CStr(.Cells(6, 2).value).Trim
                    End If
                    If Not IsNothing(.Cells(7, 2).value) Then
                        result(2) = CStr(.Cells(7, 2).value).Trim
                    End If

                End With
            End If
        Catch ex As Exception

        End Try

        getPortfolioNames = result

    End Function


    Public Function getOverloadParams() As Double()

        Dim blattName As String = "Parameters"
        Dim result As Double()
        ReDim result(1)
        result(0) = 1.05
        result(1) = 1.0

        Try
            Dim currentWB As xlns.Workbook = CType(Module1.appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                With currentWS
                    result(0) = CDbl(.Cells(1, 2).value)
                    result(1) = CDbl(.Cells(2, 2).value)
                End With
            End If
        Catch ex As Exception
            result(0) = 1.05
            result(1) = 1.0
        End Try

        getOverloadParams = result

    End Function



    ''' <summary>
    ''' retrieves the Portfolio Variant Definitions 
    ''' </summary>
    ''' <param name="kennung"></param>
    ''' <param name="blattname"></param>
    ''' <returns></returns>
    Public Function getPortfolioDefinitions(ByVal kennung As PTRpa,
                               Optional ByVal blattname As String = "") As clsPortfolioDefinitions

        Dim result As New clsPortfolioDefinitions

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1

        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattname = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattname),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            Dim lastRow As Integer
            Dim firstZeile As xlns.Range


            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    firstZeile = CType(.Rows(1), xlns.Range)
                    lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    Dim portfolioZeile As Integer = 2

                    ' now read the portfolio definitions, if there are any 

                    If kennung = PTRpa.visboCreateHedgedVariant Then

                        Dim firstPFColumn As Integer = spalte + 4
                        Dim currentColumn As Integer = firstPFColumn

                        If Not IsNothing(CType(.Cells(1, firstPFColumn), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                            Dim myPortfolioVariantName As String = CStr(CType(.Cells(1, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                            If myPortfolioVariantName <> "" Then

                                Do While myPortfolioVariantName <> ""

                                    Dim abbruch As Boolean = False
                                    Dim myPName As String = ""
                                    Dim myVName As String = ""
                                    Dim myUniqueList As New Collection
                                    Dim myPortfolioList As New List(Of String)

                                    zeile = 2

                                    Do While Not abbruch

                                        Try
                                            myPName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                            myVName = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                        Catch ex As Exception
                                            myVName = ""
                                        End Try



                                        ' now check whether or not there is a 'x' in myCurrentColumn
                                        If Not IsNothing(CType(.Cells(zeile, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                            Dim signal As String = CStr(CType(.Cells(zeile, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                                            If signal.ToLower = "x" Then

                                                If Not myUniqueList.Contains(myPName) Then
                                                    myUniqueList.Add(myPName, myPName)
                                                    Dim Key As String = calcProjektKey(myPName, myVName)
                                                    myPortfolioList.Add(Key)
                                                End If

                                            End If
                                        End If

                                        zeile = zeile + 1
                                        abbruch = (zeile > lastRow)
                                    Loop

                                    ' now add Portfolio definition 

                                    If Not result.contains(myPortfolioVariantName) Then
                                        result.addPortfolio(myPortfolioVariantName, myPortfolioList)
                                    End If


                                    ' now consider next column ...
                                    currentColumn = currentColumn + 1
                                    myPortfolioVariantName = CStr(CType(.Cells(1, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                Loop

                            End If


                        End If


                    End If

                End With

            End If

        Catch ex As Exception

        End Try

        getPortfolioDefinitions = result
    End Function


    Public Function getRanking(ByVal kennung As PTRpa,
                               Optional ByVal blattname As String = "") As SortedList(Of Integer, clsRankingParameters)

        Dim result As New SortedList(Of Integer, clsRankingParameters)

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1

        Dim projectName As String = ""
        Dim projectVariantName As String = ""

        Dim earliestStart As Date = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(1)
        Dim latestEnd As Date = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(13)


        ' given in Percentage
        Dim shortestDuration As Double = 1.0
        Dim longestDuration As Double = 1.0

        Dim aktDateTime As Date = Date.Now
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If


        ' the hedge Factorm 
        Dim hedgeFactor As Double = 1.0

        Dim lastRow As Integer
        Dim firstZeile As xlns.Range



        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattname = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattname),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    firstZeile = CType(.Rows(1), xlns.Range)
                    lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    Dim portfolioZeile As Integer = 2

                    While zeile <= lastRow

                        Dim myCurrentParams As New clsRankingParameters

                        myCurrentParams.projectName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        myCurrentParams.projectVariantName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)

                        Select Case kennung

                            Case PTRpa.visboCreateHedgedVariant

                                Try
                                    Dim tmpHedgeFactor As Double = CDbl(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If tmpHedgeFactor >= 0 And tmpHedgeFactor <= 1.0 Then
                                        myCurrentParams.hedgeFactor = tmpHedgeFactor
                                    Else
                                        myCurrentParams.hedgeFactor = 0.0
                                    End If
                                Catch ex As Exception
                                    myCurrentParams.hedgeFactor = 0.0
                                End Try

                                Try
                                    Dim tmpStartDate As Date = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If DateDiff(DateInterval.Day, tmpStartDate, Date.Now) > 0 Then
                                        myCurrentParams.newStartDate = Date.MinValue
                                    Else
                                        myCurrentParams.newStartDate = tmpStartDate
                                    End If
                                Catch ex As Exception
                                    myCurrentParams.newStartDate = Date.MinValue
                                End Try


                            Case PTRpa.visboFindProjectStart

                                myCurrentParams.earliestStart = CDate(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.latestEnd = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.shortestDuration = CDbl(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.longestDuration = CDbl(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)


                            Case PTRpa.visboFindProjectStartPM

                                myCurrentParams.earliestStart = CDate(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.latestEnd = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.shortestDuration = CDbl(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.longestDuration = CDbl(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                            Case PTRpa.visboSuggestResourceAllocation

                                Dim peopleIDs As New SortedList(Of String, Double)
                                Dim lastColumn As Integer = CType(.Cells(zeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlToLeft).Column

                                For i As Integer = 3 To lastColumn
                                    If Not IsNothing(CType(.Cells(zeile, i), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                        Dim myName As String = CStr(CType(.Cells(zeile, i), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                        Dim teamID As Integer = -1

                                        If RoleDefinitions.containsNameOrID(myName) Then
                                            Dim myRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(myName, teamID)
                                            Dim myNameID As String = RoleDefinitions.bestimmeRoleNameID(myRole.UID, teamID)

                                            If Not peopleIDs.ContainsKey(myNameID) Then
                                                peopleIDs.Add(myNameID, 0)
                                            End If

                                        End If
                                    End If
                                Next

                                myCurrentParams.peopleSuggestions = peopleIDs

                            Case Else

                        End Select

                        result.Add(zeile, myCurrentParams)

                        zeile = zeile + 1


                    End While


                End With
            End If

        Catch ex As Exception

            Throw New Exception("Fehler In Portfolio-Datei" & ex.Message)
        End Try

        getRanking = result

    End Function

    ''' <summary>
    ''' returns the sequence of the project-names 
    ''' there is only one project-variant per ranking allowed
    ''' </summary>
    ''' <returns></returns>
    Public Function getRanking(Optional ByVal blattname As String = "") As SortedList(Of Integer, String)

        Dim rankingList As New SortedList(Of Integer, String)
        Dim nameList As New SortedList(Of String, String)
        Dim key As String

        Dim zeile As Integer, spalte As Integer


        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""

        Dim lastRow As Integer


        Dim geleseneProjekte As Integer


        Dim firstZeile As xlns.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0




        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattname = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattname),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    firstZeile = CType(.Rows(1), xlns.Range)
                    lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    Dim portfolioZeile As Integer = 2

                    While zeile <= lastRow


                        pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        variantName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)

                        key = calcProjektKey(pName, variantName)

                        If Not nameList.ContainsKey(pName) Then
                            nameList.Add(pName, key)
                            If Not rankingList.ContainsKey(zeile) Then
                                rankingList.Add(zeile, key)
                            End If
                        End If

                        zeile = zeile + 1


                    End While


                End With
            End If

        Catch ex As Exception

            Throw New Exception("Fehler In Portfolio-Datei" & ex.Message)
        End Try

        getRanking = rankingList
    End Function


    Public Function readListIntoStorage(ByVal kennung As PTRpa) As Boolean
        Dim allOk As Boolean = False
        ' jetzt alle Projekte aus der Liste holen und die OverloadParams holen 

        Try
            Call logger(ptErrLevel.logInfo, "start Processing: " & kennung.ToString, "Read Projects")

            Dim listOfProjs As SortedList(Of Integer, String) = getRanking()

            For Each kvp As KeyValuePair(Of Integer, String) In listOfProjs

                Dim pname As String = getPnameFromKey(kvp.Value)
                Dim vname As String = getVariantnameFromKey(kvp.Value)
                Dim today As Date = Date.Now
                Dim hproj As clsProjekt = getProjektFromSessionOrDB(pname, vname, AlleProjekte, today)

                If Not IsNothing(hproj) Then
                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                End If

            Next
            allOk = True

        Catch ex As Exception
            allOk = False
        End Try
        readListIntoStorage = allOk

    End Function


    Public Function processFindProjectStart(ByVal myName As String, Optional ByVal myKennung As PTRpa = PTRpa.visboFindProjectStart) As Boolean

        Dim allOk As Boolean = False

        Try
            Dim jobParameters As clsJobParameters = getJobParameters("Parameters", myKennung)

            If jobParameters.portfolioName = "" Then
                jobParameters.portfolioName = myActivePortfolio
            End If

            Dim portfolioName As String = jobParameters.portfolioName

            Dim aggregationList As New List(Of String)
            Dim skillList As New List(Of String)


            If myKennung = PTRpa.visboFindProjectStart Then
                Call logger(ptErrLevel.logInfo, "start Processing find best Start with regard to roles & skills: ", myName)
            Else
                Call logger(ptErrLevel.logInfo, "start Processing Find best Start with regard to Phases , Milestones: ", myName)
            End If


            Dim readProjects As Integer = 0
            Dim createdProjects As Integer = 0
            'Dim importedProjects As Integer = ImportProjekte.Count

            Dim outPutCollection As New Collection



            ' jetzt alle Projekte aus der Liste holen und die OverloadParams holen 
            Try
                'Dim listOfProjs As SortedList(Of Integer, String) = getRanking()
                Dim listOfProjs As SortedList(Of Integer, clsRankingParameters) = getRanking(myKennung)

                If listOfProjs.Count > 0 Then
                    allOk = True
                Else
                    Dim msgTxt As String = "no new project names were given "
                    outPutCollection.Add(msgTxt)
                    allOk = False
                End If
                For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In listOfProjs


                    Try
                        Dim pname As String = kvp.Value.projectName
                        Dim vname As String = kvp.Value.projectVariantName
                        Dim today As Date = Date.Now
                        Dim hproj As clsProjekt = getProjektFromSessionOrDB(pname, vname, AlleProjekte, today)

                        If Not IsNothing(hproj) Then
                            ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                            allOk = allOk And True
                        Else
                            Dim msgTxt As String = "could not find " & hproj.getShapeText
                            outPutCollection.Add(msgTxt)
                            allOk = False
                        End If
                    Catch ex As Exception
                        allOk = False
                    End Try

                    If Not allOk Then
                        Exit For
                    End If
                Next

            Catch ex As Exception
                allOk = False
            End Try


            If allOk Then
                Call logger(ptErrLevel.logInfo, "Project List imported: " & myName, ImportProjekte.Count & " read; ")
            Else
                Call logger(ptErrLevel.logError, "failure in reading new projects: " & myName, outPutCollection)
            End If

            If allOk Then

                Dim noActivePortfolio As Boolean = True
                Dim dbPortfolioNames As New SortedList(Of String, String)

                ' if Portfolio with active Projects is given and exists:  
                ' then we probably do have a brownfield
                If portfolioName <> "" Then

                    Dim errMsg As New clsErrorCodeMsg
                    dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)
                    noActivePortfolio = Not dbPortfolioNames.ContainsKey(portfolioName)
                End If

                If noActivePortfolio Then
                    Call logger(ptErrLevel.logError, "no active Portfolio: " & portfolioName, myKennung.ToString)
                Else
                    ' check whether and how projects are fitting to the already existing Portfolio 
                    allOk = processProjectListWithActivePortfolio(jobParameters, myKennung)

                End If

            Else
                ' no additional logger necessary - is done in storeImportProjekte
            End If


            ' now empty the complete session  
            Call emptyRPASession()
            Call logger(ptErrLevel.logInfo, "end Processing: " & myKennung.ToString, myName)

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "errors occurred when processing: " & myKennung.ToString, myName & ": " & ex.Message)
        End Try

        processFindProjectStart = allOk

    End Function


    ''' <summary>
    ''' creates hedged variants for existing projects
    ''' projects need to be imported already with readListIntoStorage
    ''' </summary>
    ''' <returns></returns>
    Public Function processCreateHedgedVariants() As Boolean

        Dim result As Boolean = False
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""


        Try

            ' now get Portfolio Name and Variante-NAme 
            Dim params() As String = getPortfolioNames()
            Dim portfolioName As String = params(0)
            Dim portfolioVariantName As String = params(1)

            If portfolioName = "" Then
                portfolioName = myActivePortfolio
            End If

            Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(PTRpa.visboCreateHedgedVariant)

            Dim projectVariantName As String = "hedged"
            If params(2) <> "" Then
                projectVariantName = params(2).Trim
            End If

            Dim outPutCollection As New Collection

            For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In rankingList

                Dim projectstoreRequired As Boolean = False
                Dim variantStoreRequired As Boolean = False

                Dim key As String = calcProjektKey(kvp.Value.projectName, kvp.Value.projectVariantName)
                Dim hproj As clsProjekt = ImportProjekte.getProject(key)
                If Not IsNothing(hproj) Then
                    ' now move the project 
                    If hproj.vpStatus = "initial" Or hproj.vpStatus = "initialized" Or hproj.vpStatus = "vorgeschlagen" Or hproj.vpStatus = "proposed" Then

                        Dim newStartDate As Date = kvp.Value.newStartDate
                        If DateDiff(DateInterval.Day, Date.Now, newStartDate) >= 0 Then

                            Dim deltaInDays As Integer = DateDiff(DateInterval.Day, hproj.startDate, newStartDate)

                            If deltaInDays <> 0 Then
                                Dim newEndDate As Date = hproj.endeDate.AddDays(deltaInDays)
                                ' try to set it to movable, will only be successful if conditions are met ...
                                hproj.movable = True
                                If hproj.movable Then
                                    Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                                    If Not IsNothing(tmpProj) Then
                                        If (DateDiff(DateInterval.Day, tmpProj.startDate, newStartDate) + DateDiff(DateInterval.Day, tmpProj.endeDate, newEndDate)) = 0 Then
                                            hproj = tmpProj
                                            projectstoreRequired = True
                                            msgTxt = "project moved to: " & newStartDate.ToShortDateString & " - " & newEndDate.ToShortDateString
                                            Call logger(ptErrLevel.logInfo, "process Create Hedged Variants, ", msgTxt)
                                        Else
                                            msgTxt = "project could not be moved correctly: " & hproj.getShapeText
                                            Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                                        End If
                                    Else
                                        msgTxt = "project could not be moved at all"
                                    End If
                                Else
                                    If hproj.hasActualValues Then
                                        msgTxt = "Project with actual data can not be moved " & hproj.getShapeText
                                    Else
                                        msgTxt = "Project with Status other than initial or proposed can not be moved " & hproj.getShapeText
                                    End If

                                    Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                                End If

                            End If

                            ' now create the variant with appropriate hedgeFactor 
                            Dim variantProj As clsProjekt = Nothing
                            If kvp.Value.hedgeFactor > 0 And kvp.Value.hedgeFactor < 1 Then
                                variantProj = hproj.createHedgedVariant(kvp.Value.hedgeFactor)
                                If Not IsNothing(variantProj) Then
                                    variantStoreRequired = True
                                End If
                            End If


                            If projectstoreRequired Then
                                outPutCollection.Clear()

                                msgTxt = hproj.getShapeText & " : " & hproj.startDate.ToShortDateString

                                ' make sure it is in AlleProjekte, becaue store Method requires it being in AlleProjekte 
                                Dim didExist As Boolean = AlleProjekte.Containskey(calcProjektKey(hproj))
                                If Not didExist Then
                                    AlleProjekte.Add(hproj, False)
                                End If

                                If storeSingleProjectToDB(hproj, outPutCollection) Then
                                    Call logger(ptErrLevel.logInfo, "project with new startDate stored: ", msgTxt)
                                    'Console.WriteLine("project with new startDate stored: " & msgTxt)
                                    'If Not setWriteProtection(hproj, False) Then
                                    '    Call logger(ptErrLevel.logWarning, "Aufheben Write PRotection did not work ...  ", hproj.getShapeText)
                                    'End If
                                Else
                                    Call logger(ptErrLevel.logError, "project store with new startDate failed: " & msgTxt, outPutCollection)
                                    'Console.WriteLine("!! ... project store with new startDate failed: " & msgTxt)
                                End If

                                If Not didExist Then
                                    AlleProjekte.Remove(calcProjektKey(hproj), False)
                                End If

                            End If

                            If variantStoreRequired Then
                                outPutCollection.Clear()

                                msgTxt = variantProj.getShapeText & " : " & variantProj.variantDescription & vbLf & " hedgeFactor: " & kvp.Value.hedgeFactor * 100 & "%"

                                Dim didExist As Boolean = AlleProjekte.Containskey(calcProjektKey(variantProj))
                                If Not didExist Then
                                    AlleProjekte.Add(variantProj, False)
                                End If
                                If storeSingleProjectToDB(variantProj, outPutCollection) Then
                                    Call logger(ptErrLevel.logInfo, "hedged variant stored: ", msgTxt)
                                Else
                                    Call logger(ptErrLevel.logError, "hedged variant store failed: " & msgTxt, outPutCollection)
                                    'Console.WriteLine("!! ... hedged variant store failed: " & msgTxt)
                                End If

                                If Not didExist Then
                                    AlleProjekte.Remove(calcProjektKey(variantProj), False)
                                End If

                            End If
                        Else
                            msgTxt = "not a appropriate new Startdate : " & hproj.getShapeText & newStartDate.ToShortDateString
                            Call logger(ptErrLevel.logInfo, "process Create Hedged Variants, ", msgTxt)
                        End If

                    Else
                        msgTxt = "Project with Status other than initial or proposed can not be moved " & hproj.getShapeText
                        Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                    End If
                Else
                    msgTxt = "Project does not exist " & kvp.Value.projectName
                    Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                End If

            Next

            ' now read the new portfolios 
            Dim myPortfolioVariants As clsPortfolioDefinitions = getPortfolioDefinitions(PTRpa.visboCreateHedgedVariant)

            If myPortfolioVariants.portfolioListe.Count > 0 Then

                ' get all projects of Active Portfolio , put them into AlleActiveProjects
                Dim activePortfolioProjects As New clsProjekteAlle
                If putPortfolioIntoSession(myActivePortfolio, "", activePortfolioProjects) Then

                    ' 2: now for each portfolio Variant : 
                    ' get all projects of AlleActiveProjects , put them into AlleProjekte and in ShowProjekte 



                    For Each kvp As KeyValuePair(Of String, List(Of String)) In myPortfolioVariants.portfolioListe

                        Try
                            AlleProjekte.Clear()
                            ShowProjekte.Clear()

                            For Each activeKVP As KeyValuePair(Of String, clsProjekt) In activePortfolioProjects.liste

                                AlleProjekte.Add(activeKVP.Value)
                                If Not ShowProjekte.contains(activeKVP.Value.name) Then
                                    ShowProjekte.Add(activeKVP.Value)
                                End If

                            Next

                            ' get all projects referenced in a PortfolioVariantList , put them into AlleProjekte , ShowProjekte 
                            For Each tmpKey As String In kvp.Value
                                Dim hproj As clsProjekt = ImportProjekte.getProject(tmpKey)
                                If Not IsNothing(hproj) Then
                                    AlleProjekte.Add(hproj)
                                    ShowProjekte.AddAnyway(hproj)
                                End If
                            Next

                            ' create the Portfolio and store it 
                            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                         cName:=portfolioName, vName:=kvp.Key)

                            outPutCollection.Clear()
                            Call storeSingleConstellationToDB(outPutCollection, toStoreConstellation, Nothing)

                            msgTxt = toStoreConstellation.constellationName & " ( " & toStoreConstellation.variantName & " )"

                            'Console.WriteLine("Portfolio Variant stored: " & msgTxt)
                            Call logger(ptErrLevel.logInfo, "Portfolio Variant stored: ", msgTxt)


                        Catch ex As Exception
                            Call logger(ptErrLevel.logError, "Failure when preparing store of portfolio ", ex.Message)
                        End Try


                    Next

                    result = True

                End If


            Else
                result = True
                Call logger(ptErrLevel.logInfo, "no Portfolio Variants created because no Portfolio Name for results was provided", "")
            End If



        Catch ex As Exception

        End Try


        ' now empty the complete session  
        Call emptyRPASession()

        processCreateHedgedVariants = result

    End Function

    Public Function processAutoAllocatePortfolio() As Boolean

        Dim result As Boolean = True
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""

        Dim atleastOneError As Boolean = False

        Dim outputCollection As New Collection

        Dim heute As Date = Date.Now

        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        ' set it back to undefined
        showRangeLeft = 0
        showRangeRight = 0

        Try

            ' now get Portfolio Name and Variante-NAme 
            Dim params() As String = getPortfolioNames()

            Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(PTRpa.visboSuggestResourceAllocation)


            Dim portfolioName As String = params(0)
            Dim variantName As String = params(1)

            Dim projectVariantName As String = "auto"
            If params(2) <> "" Then
                projectVariantName = params(2).Trim
            End If

            Dim autoAllocate As Boolean = True

            ShowProjekte.Clear()
            AlleProjekte.Clear()
            ImportProjekte.Clear()
            projectConstellations.Clear()

            ' now load the the portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(portfolioName,
                                                                                               "", cTime, Err, variantName:=variantName, storedAtOrBefore:=heute)


            If Not IsNothing(myConstellation) Then

                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen

                projectConstellations.Add(myConstellation)

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)

                    If Not IsNothing(hproj) Then

                        AlleProjekte.Add(hproj)
                        ShowProjekte.AddAnyway(hproj)

                    Else
                        Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                        'Console.WriteLine("Loading " & kvp.Key & " failed ..", " Operation continued ...")
                        atleastOneError = True
                    End If

                Next

                ' now do the operation 

                For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In rankingList

                    Dim myProj As clsProjekt = ShowProjekte.getProject(kvp.Value.projectName)
                    Dim fmsg As String = ""


                    ' now Create a Variant from that , if it is not already the very same variant
                    If myProj.variantName <> projectVariantName Then
                        myProj = myProj.createVariant(projectVariantName, "auto-created variant")

                        ' now put it into ShowProjekte , AlleProjekte
                        ShowProjekte.AddAnyway(myProj)
                        AlleProjekte.Add(myProj)
                    End If


                    If Not IsNothing(myProj) Then

                        Call ShowProjekte.autoAllocate(myProj.name, "", False, fmsg, suggestedIDs:=kvp.Value.peopleSuggestions)

                        If fmsg = "" Then

                            outputCollection.Clear()
                            Dim storeProj As clsProjekt = ShowProjekte.getProject(myProj.name)

                            If Not IsNothing(storeProj) Then
                                If Not atleastOneError Then
                                    If storeSingleProjectToDB(storeProj, outputCollection) Then
                                        Call logger(ptErrLevel.logInfo, "project team allocated and stored: ", storeProj.getShapeText)
                                        'Console.WriteLine("project team allocated and stored: " & storeProj.getShapeText)

                                        'If Not setWriteProtection(storeProj, False) Then
                                        '    Call logger(ptErrLevel.logWarning, "Aufheben Write PRotection did not work ...  ", storeProj.getShapeText)
                                        'End If
                                    Else
                                        Call logger(ptErrLevel.logError, "store project failed : " & storeProj.getShapeText, outputCollection)
                                        'Console.WriteLine("!! ... store project team allocation failed : " & storeProj.getShapeText)
                                        atleastOneError = True
                                    End If
                                Else
                                    Call logger(ptErrLevel.logInfo, "because former Error occurred , no store happened : " & storeProj.getShapeText, outputCollection)
                                    'Console.WriteLine("!! ... because former Error occurred , no store happened: " & storeProj.getShapeText)
                                    atleastOneError = True
                                End If

                            End If

                        Else
                            ' failure 
                            atleastOneError = True
                            Call logger(ptErrLevel.logError, "Auto-Allocation failure: " & kvp.Key & " " & fmsg, " ... Operation continued ...")
                            'Console.WriteLine("!! ... Auto-Allocation failure : " & myProj.getShapeText)
                        End If
                    Else
                        ' failure 
                        atleastOneError = True
                        Call logger(ptErrLevel.logError, "Auto-Allocation failure: could not read " & kvp.Value.projectName & " " & kvp.Value.projectVariantName, " ... Operation continued ...")
                        'Console.WriteLine("Auto-Allocation failure: could not read " & kvp.Value.projectName & " " & kvp.Value.projectVariantName)

                    End If


                Next

                Call logger(ptErrLevel.logInfo, "Auto-Allocating Projects from Portfolio " & myActivePortfolio, " End of Operation ... ")

                ' now create the according Portfolio


            Else
                atleastOneError = True
                msgTxt = "Load Portfolio " & myActivePortfolio & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                'Console.WriteLine(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio & " failed ..")
                Throw New ArgumentException(msgTxt)
            End If

            If Not atleastOneError Then

                Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=projectVariantName)

                outputCollection.Clear()
                Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                If outputCollection.Count > 0 Then
                    Call logger(ptErrLevel.logInfo, "Project List with Active Portfolio: ", outputCollection)
                    'Console.WriteLine("Portfolio created " & portfolioName & " ( " & projectVariantName & " )")
                End If

            Else
                Call logger(ptErrLevel.logInfo, "no store of portfolio because of former Error .. ", "")
                'Console.WriteLine("no store of portfolio because of former Error .. ")
            End If


        Catch ex As Exception

        End Try

        ' restore values 
        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        Call emptyRPASession()

        processAutoAllocatePortfolio = result

    End Function

    ''' <summary>
    ''' create a Portfolio Variant and according project variants in a way that there are no more any bottlenecks at people base
    ''' all projects of the ranking list should be handeled, Sequence represents the priority 
    ''' all projects of ranking list are first excluded from ShowProjekte , then taken in one by one according priority  
    ''' all other projects of portfolio whcih are not listed in the ranking list remain unchanged
    ''' in ranking list there may be new projects not yet contained in given portfolio 
    ''' i.e values distributed so that no person is being overloaded, 
    ''' values geing beyond that are assigned the according summary role and in a second step being auto-Allocated 
    ''' </summary>
    ''' <returns></returns>
    Public Function processAutoAdjustPortfolio() As Boolean

        Dim result As Boolean = True
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""



        Dim outputCollection As New Collection

        Dim heute As Date = Date.Now

        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        ' set it back to undefined
        showRangeLeft = 0
        showRangeRight = 0

        Try

            ' now get Portfolio Name and Variante-NAme 
            'Dim params() As String = getPortfolioNames()

            ' tk 25.6.22 
            Dim myKennung As PTRpa = PTRpa.visboAutoAdjust
            Dim jobParameters As clsJobParameters = getJobParameters("Parameters", myKennung)
            'Dim exceptionList As Collection = getNameList("Exception List")

            'Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(myKennung)
            Dim ranking As New clsRankingList With {
                .liste = getRanking(myKennung)
            }

            'Dim portfolioName As String = params(0)
            Dim portfolioName As String = jobParameters.portfolioName
            'Dim variantName As String = params(1)
            Dim portfolioVariantName As String = jobParameters.portfolioVariantName

            Dim projectVariantName As String = "auto"
            'If params(2) <> "" Then
            If jobParameters.projectVariantName <> "" Then
                projectVariantName = jobParameters.projectVariantName.Trim
            End If

            Dim autoAllocate As Boolean = True

            ShowProjekte.Clear()
            AlleProjekte.Clear()
            projectConstellations.Clear()

            ' now load the portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(portfolioName,
                                                                                               "", cTime, Err, variantName:="", storedAtOrBefore:=heute)


            If Not IsNothing(myConstellation) Then

                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & portfolioName, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen

                projectConstellations.Add(myConstellation)

                ' now get each project into session, except those which need to be changed according ranking
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)

                    If Not ranking.containsPName(pName) Then

                        Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)
                        If Not IsNothing(hproj) Then
                            AlleProjekte.Add(hproj)

                            ' if it is already in ShowProjekte: remove it , then add this one 
                            ShowProjekte.AddAnyway(hproj)
                        Else
                            Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                        End If

                    End If

                Next

                ' now all projects of portfolio which shoul not be changed but considered n context are loaded ... 

                For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In ranking.liste
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(kvp.Value.projectName, kvp.Value.projectVariantName, AlleProjekte, heute)

                    If Not IsNothing(hproj) Then

                        If hproj.variantName <> projectVariantName Then
                            hproj = hproj.createVariant(projectVariantName, "auto-created variant")
                        End If

                        AlleProjekte.Add(hproj)

                        ' if it is already in ShowProjekte: remove it , then add this one 
                        ShowProjekte.AddAnyway(hproj)

                        Dim fmsg As String = ""
                        Call ShowProjekte.autoDistribute(hproj.name, "", fmsg)

                        If fmsg = "" Then
                            ' success
                            Call logger(ptErrLevel.logInfo, "Adjustment successful: " & kvp.Key, " ... Operation continued ...")

                            Call ShowProjekte.autoAllocate(hproj.name, "", True, fmsg)

                            If fmsg = "" Then

                                outputCollection.Clear()

                                If storeSingleProjectToDB(hproj, outputCollection) Then
                                    Call logger(ptErrLevel.logInfo, "project variant adjusted and stored: ", hproj.getShapeText)

                                Else
                                    Call logger(ptErrLevel.logError, "project variant store failed: " & hproj.getShapeText, outputCollection)
                                    'Console.WriteLine("!! ... project store failed: " & kvp.Value.getShapeText)
                                End If

                                Call logger(ptErrLevel.logInfo, "Auto-Allocation successful: " & hproj.name, " ... Operation continued ...")
                            Else
                                ' failure 
                                Call logger(ptErrLevel.logError, "Auto-Allocation failure: " & hproj.name & " " & fmsg, " ... Operation continued ...")
                            End If
                        Else
                            ' failure 
                            Call logger(ptErrLevel.logError, "Adjustment failure: " & hproj.name & " " & fmsg, " ... Operation continued ...")
                        End If

                    Else
                        Call logger(ptErrLevel.logWarning, "Loading of Rank " & kvp.Key & ": " & kvp.Value.projectName & " failed ..", " Operation continued ...")
                    End If
                Next


                Call logger(ptErrLevel.logInfo, "Adjusting Projects from Portfolio " & portfolioName, " End of Operation ... ")

                ' now create the according Portfolio


            Else
                msgTxt = "Load Portfolio " & portfolioName & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & portfolioName, " failed ..")
                Throw New ArgumentException(msgTxt)
            End If

            If portfolioVariantName <> "" Then
                Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=portfolioVariantName)

                outputCollection.Clear()
                Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                If outputCollection.Count > 0 Then
                    Call logger(ptErrLevel.logInfo, "Portfolio Variant stored: ", outputCollection)
                End If
            End If


        Catch ex As Exception

        End Try




        ' restore values 
        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        Call emptyRPASession()

        processAutoAdjustPortfolio = result
    End Function


    ''' <summary>
    ''' reads all projects of a given Portfolio into storage: all projects are then within clsProjekteAlle 'sessionListe'
    ''' updateCurrentConstellation is by default set to false, i.e a currentSessionConstellation is not defined by that in Default. 
    ''' </summary>
    ''' <param name="myPortfolioName"></param>
    ''' <param name="myPortfolioVName"></param>
    ''' <param name="sessionListe"></param>
    ''' <returns></returns>
    Public Function putPortfolioIntoSession(ByVal myPortfolioName As String, ByVal myPortfolioVName As String, ByRef sessionListe As clsProjekteAlle,
                                             Optional ByVal upDateCurrentConstellation As Boolean = False) As Boolean

        Dim allOk As Boolean = False
        Dim Err As New clsErrorCodeMsg


        Try
            sessionListe.Clear(upDateCurrentConstellation)

            If myPortfolioName = "" Then
                myPortfolioName = myActivePortfolio
                myPortfolioVName = ""
            End If

            Dim heute As Date = Date.Now
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(myPortfolioName,
                                                                                           "", heute, Err, variantName:=myPortfolioVName, storedAtOrBefore:=heute)


            If Not IsNothing(myConstellation) Then

                ' tmpname in die Session-Liste wieder aufnehmen
                projectConstellations.Add(myConstellation)
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, sessionListe, heute)
                    If Not IsNothing(hproj) Then

                        sessionListe.Add(hproj, upDateCurrentConstellation)

                    End If
                Next

                allOk = True
            Else
                Dim msgTxt As String = "Load Portfolio " & myPortfolioName & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                allOk = False
            End If

        Catch ex As Exception
            allOk = False
        End Try

        putPortfolioIntoSession = allOk

    End Function


    ''' <summary>
    ''' in ImportProjekte sind alle aktuell eingelesenen Projekte 
    ''' </summary>
    ''' <returns></returns>
    Public Function processProjectListWithActivePortfolio(ByVal jobParameters As clsJobParameters, ByVal myKennung As PTRpa) As Boolean
        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""

        Dim heute As Date = Date.Now

        ' cache löschen
        Dim result0 As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Dim myActivePortfolio As String = jobParameters.portfolioName
        Dim portfolioVariantName As String = jobParameters.portfolioVariantName

        Dim projectVariantName As String = jobParameters.projectVariantName
        If projectVariantName = "" Then
            projectVariantName = "arb"
        End If

        Dim aggregationList As New List(Of String)
        Dim skillList As New List(Of String)

        ' for showing which projects could not be considered
        Dim missingList As New clsProjekteAlle

        If myKennung = PTRpa.visboFindProjectStart Then
            ' build aggregation List
            ' now get the aggregation Roles
            Dim aggregationRoles As SortedList(Of Integer, String) = RoleDefinitions.getAggregationRoles()
            Dim teamID As Integer = -1


            ' currently only Exclude of Roles & Skills is supported ..
            ' checkout aggregation Roles
            For Each ar As KeyValuePair(Of Integer, String) In aggregationRoles
                Dim tmpStrID As String = RoleDefinitions.bestimmeRoleNameID(ar.Key, teamID)
                If Not aggregationList.Contains(tmpStrID) Then
                    If jobParameters.donotConsiderRoleSkills.Count = 0 Then
                        aggregationList.Add(tmpStrID)
                    Else
                        If Not jobParameters.donotConsiderRoleSkills.Contains(ar.Value) Then
                            aggregationList.Add(tmpStrID)
                        End If
                    End If
                End If
            Next

            ' build Skill List 
            Dim skillIDs As Collection = ImportProjekte.getRoleSkillIDs()


            For Each si As String In skillIDs

                ' new  
                If Not skillList.Contains(si) Then
                    If jobParameters.donotConsiderRoleSkills.Count = 0 Then
                        skillList.Add(si)
                    Else
                        If Not jobParameters.donotConsiderRoleSkills.Contains(si) Then
                            skillList.Add(si)
                        End If
                    End If
                End If
            Next
        End If



        Try
            ShowProjekte.Clear()
            AlleProjekte.Clear()

            ' now load the existing portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(myActivePortfolio,
                                                                                               "", cTime, Err, variantName:="", storedAtOrBefore:=heute)


            ' tk 1.5. 
            Dim nextLineNumber As Integer = myConstellation.getMaxRowNumber + 1

            If Not IsNothing(myConstellation) Then
                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen
                projectConstellations.Add(myConstellation)

                ' rowNr ist needed to use it definietion of Portfolio sorting / sequence

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)

                    Try
                        hproj.tfZeile = myConstellation.getBoardZeile(pName)
                        If hproj.tfZeile > nextLineNumber Then
                            nextLineNumber = hproj.tfZeile + 1
                        End If
                    Catch ex As Exception
                        hproj.tfZeile = 2
                    End Try

                    If Not IsNothing(hproj) Then

                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                        ' removes hproj from ShowProjekte, if already in there
                        ShowProjekte.AddAnyway(hproj)

                    Else
                        Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                    End If

                Next

                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " End of Operation ... ")

            Else
                msgTxt = "Load Portfolio " & myActivePortfolio & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                Throw New ArgumentException(msgTxt)
            End If

            ' get the ranking list 
            'Dim rankingList As SortedList(Of Integer, String) = getRanking()
            Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(myKennung)


            ' now create a Portfolio variant with unchanged new projects ...
            Dim removeSPList As New List(Of String)
            Dim removeAPList As New List(Of String)

            Dim first As Boolean = True
            Dim minMonthColumn As Integer = 0
            Dim maxMonthColumn As Integer = 0

            Dim outputCollection As New Collection

            Dim myRowNr As Integer = nextLineNumber
            For Each rankingPair As KeyValuePair(Of Integer, clsRankingParameters) In rankingList
                Dim key As String = calcProjektKey(rankingPair.Value.projectName, rankingPair.Value.projectVariantName)
                Dim hproj As clsProjekt = ImportProjekte.getProject(key)
                If Not IsNothing(hproj) Then

                    ' bestimme die Line Number 
                    hproj.tfZeile = myRowNr

                    If first Then
                        first = False

                        minMonthColumn = System.Math.Min(getColumnOfDate(hproj.startDate), getColumnOfDate(rankingPair.Value.earliestStart))
                        maxMonthColumn = getColumnOfDate(hproj.endeDate)
                    Else
                        'Dim myMin As Integer = getColumnOfDate(hproj.startDate)
                        Dim myMin As Integer = System.Math.Min(getColumnOfDate(hproj.startDate), getColumnOfDate(rankingPair.Value.earliestStart))
                        Dim myMax As Integer = getColumnOfDate(hproj.endeDate)
                        If myMin < minMonthColumn Then
                            minMonthColumn = myMin
                        End If
                        If myMax > maxMonthColumn Then
                            maxMonthColumn = myMax
                        End If
                    End If

                    hproj.movable = True
                    ' check whether or not project is beginning after today ..
                    'If DateDiff(DateInterval.Day, Date.Now, hproj.startDate) < 0 Then
                    If Not hproj.movable Then
                        outputCollection.Add(hproj.getShapeText)
                    End If

                    If Not AlleProjekte.Containskey(key) Then
                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                        removeAPList.Add(key)
                    Else
                        ' bring updated hproj into AlleProjekte
                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                    End If

                    If Not ShowProjekte.contains(hproj.name) Then
                        ShowProjekte.Add(hproj)
                        removeSPList.Add(hproj.name)
                    Else
                        ' do nothing - it is already in there
                        ' at least as a variant  
                        'ShowProjekte.AddAnyway(hproj)
                    End If


                End If

                myRowNr = myRowNr + 1
            Next

            ' now Check whether or not minMonthCol ist in Future, if not end it , because that is not allowed 
            If minMonthColumn < getColumnOfDate(Date.Now) Then
                Call logger(ptErrLevel.logError, "Find best Project Start - no projects allowed to start today or before today ", outputCollection)
                result = False
                ShowProjekte.Clear()
                AlleProjekte.Clear()
                ImportProjekte.Clear(False)
            End If

            If result = True Then


                Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                cName:=myActivePortfolio, vName:=portfolioVariantName)


                Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                If outputCollection.Count > 0 Then
                    Call logger(ptErrLevel.logInfo, "Project List Import, Store Portfolio-Variant " & portfolioVariantName & " result:", outputCollection)
                End If

                ' now rest Showprojekte to formerStatus 
                For Each tmpName As String In removeAPList
                    AlleProjekte.Remove(tmpName)
                Next

                For Each tmpName As String In removeSPList
                    ShowProjekte.Remove(tmpName)
                Next

                ' nach dem Remove ist evtl wieder das SortCriteria zurückgesetzt
                ' die Sortierung wieder nach Zeile herstellen 
                currentSessionConstellation.sortCriteria = ptSortCriteria.customTF


                ' now check whether there are overutilizations 
                ' if so , move showRangeLeft and showrangeRight  1 by 1 , until there are no overutilizations any more 

                ' only need to be considered from Today on
                showRangeLeft = System.Math.Max(minMonthColumn, getColumnOfDate(Date.Now))
                showRangeRight = System.Math.Max(minMonthColumn + 1, maxMonthColumn)

                Dim stopValue As Integer = showRangeRight

                Dim overutilizationFound As Boolean = False
                Dim referenceMSValues As Double() = Nothing
                Dim referencePHValues As Double() = Nothing

                Dim sumIterations As Integer = 0

                If myKennung = PTRpa.visboFindProjectStart Then
                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                End If


                If overutilizationFound And myKennung = PTRpa.visboFindProjectStart Then
                    msgTxt = "there are already resource bottlenecks in the starting portfolio " & myActivePortfolio
                    Call logger(ptErrLevel.logError, msgTxt, " please solve this first before considering new projects ... calculation stopped ..")
                    result = False
                End If
                '

                If result = True Then



                    ' create variant , if necessary
                    ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
                    ' It then represents the sequence: Row1 is the most important project 

                    myRowNr = nextLineNumber


                    '
                    '
                    ' now iterate through all the projects 
                    ' ####################################
                    For Each rankingPair As KeyValuePair(Of Integer, clsRankingParameters) In rankingList

                        sumIterations = 0

                        ' now check whether another hproj.variant is already in ShowProjekte - in ths case do nothing and 
                        ' process next element
                        If Not ShowProjekte.Liste.ContainsKey(rankingPair.Value.projectName) Then

                            Dim key As String = calcProjektKey(rankingPair.Value.projectName, rankingPair.Value.projectVariantName)
                            Dim hproj As clsProjekt = ImportProjekte.getProject(key)

                            If Not IsNothing(hproj) Then

                                Dim storeRequired As Boolean = False

                                Try
                                    hproj.tfZeile = myRowNr
                                Catch ex As Exception

                                End Try

                                ' now first check whether or not hproj is already positioned on earliest StartDate
                                ' if not then move it towards the earliest startdate
                                Dim newStartDate As Date = hproj.startDate
                                Dim newEndDate As Date = hproj.endeDate

                                Dim stdDuration As Integer = hproj.dauerInDays
                                Dim myDuration As Integer = stdDuration
                                'Dim minDuration As Integer = CInt(stdDuration * 0.7)

                                Dim minDuration As Integer = stdDuration
                                Dim maxDuration As Integer = stdDuration

                                If rankingPair.Value.shortestDuration > 5 Then
                                    minDuration = System.Math.Min(rankingPair.Value.shortestDuration, stdDuration)
                                    maxDuration = System.Math.Max(rankingPair.Value.longestDuration, stdDuration)
                                ElseIf rankingPair.Value.shortestDuration <= 1.0 Then
                                    minDuration = CInt(stdDuration * rankingPair.Value.shortestDuration)
                                    maxDuration = CInt(stdDuration * rankingPair.Value.longestDuration)
                                End If




                                Dim startOffset As Integer = 0
                                Dim durationModifier As Integer = (stdDuration - minDuration)

                                If DateDiff(DateInterval.Day, hproj.startDate, rankingPair.Value.earliestStart) <> 0 Then
                                    startOffset = DateDiff(DateInterval.Day, hproj.startDate, rankingPair.Value.earliestStart)
                                End If



                                If startOffset <> 0 Or durationModifier > 0 Then

                                    ' because now project is going to get shortened or moved resp both. 
                                    storeRequired = True

                                    ' create variant if not already done
                                    If Not hproj.variantName.EndsWith(projectVariantName) Then
                                        Dim useVariantName As String = ""
                                        If hproj.variantName <> "" Then
                                            useVariantName = hproj.variantName & " " & projectVariantName
                                        Else
                                            useVariantName = projectVariantName
                                        End If
                                        hproj = hproj.createVariant(useVariantName, "variant to avoid bottlenecks")
                                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                    End If

                                    newStartDate = hproj.startDate.AddDays(startOffset)
                                    newEndDate = hproj.endeDate.AddDays(startOffset - durationModifier)

                                    hproj.movable = True
                                    Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)


                                    If Not IsNothing(tmpProj) Then
                                        hproj = tmpProj

                                        Dim tmpMsg As String = "try out: " & hproj.getShapeText & newStartDate & " - " & newEndDate
                                        Call logger(ptErrLevel.logInfo, "status:  ", tmpMsg)

                                    End If
                                End If



                                Dim latestEndDate As Date = rankingPair.Value.latestEnd


                                ' now define showrangeLeft and showrangeRight from hproj 
                                showRangeLeft = getColumnOfDate(hproj.startDate)
                                showRangeRight = getColumnOfDate(hproj.endeDate)

                                ' have to happen here because just before hproj is added to ShowProjekte, find out what the situation is before ...
                                If myKennung = PTRpa.visboFindProjectStartPM Then
                                    ' now define the reference Values for Phases and Milestones 
                                    referenceMSValues = ShowProjekte.getMilestonesFrequency(jobParameters.getMilestoneNames)
                                    referencePHValues = ShowProjekte.getPhaseFrequency(jobParameters.getPhaseNames)
                                Else
                                    ' now define skill-List, because it is good enough to only consider skills of the hproj under consideration
                                    skillList.Clear()
                                    Dim skillIDs As Collection = hproj.getSkillNameIds

                                    For Each si As String In skillIDs
                                        If Not skillList.Contains(si) Then
                                            If jobParameters.donotConsiderRoleSkills.Count = 0 Then
                                                skillList.Add(si)
                                            Else
                                                If Not jobParameters.donotConsiderRoleSkills.Contains(si) Then
                                                    skillList.Add(si)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If


                                ' check auf Exists is not necessary with AlleProjekte, because it will be replaced if it already exists 
                                AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                ShowProjekte.AddAnyway(hproj)


                                If myKennung = PTRpa.visboFindProjectStart Then
                                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                                Else
                                    overutilizationFound = ShowProjekte.overLoadMSPhasesFound(jobParameters.getMilestoneNames, jobParameters.limitMilestones,
                                                                                                  referenceMSValues,
                                                                                                  jobParameters.getPhaseNames, jobParameters.limitPhases,
                                                                                                  referencePHValues)
                                End If



                                If overutilizationFound Then

                                    ' create variant if not already done
                                    If Not hproj.variantName.EndsWith(projectVariantName) Then
                                        Dim useVariantName As String = ""
                                        If hproj.variantName <> "" Then
                                            useVariantName = hproj.variantName & " " & projectVariantName
                                        Else
                                            useVariantName = projectVariantName
                                        End If
                                        hproj = hproj.createVariant(useVariantName, "variant to avoid bottlenecks")
                                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                    End If


                                    Dim deltaInDays As Integer = jobParameters.defaultDeltaInDays
                                    ' now modify this one ...


                                    Dim startIterations As Integer = 0
                                    Dim durationIterations As Integer = 0

                                    ' before doing the iterations trying out different lengths of projects from minDuration to maxDuration
                                    ' first try out all shortest possible durations ...

                                    Dim firstIteration As Boolean = True

                                    Dim firsthproj As clsProjekt = hproj

                                    For i = 1 To 2
                                        ' first iteration: try out different starting points and only shortest durations
                                        ' second iteration: try out different starting points and different lengths 
                                        If overutilizationFound Then

                                            Try
                                                Dim tmpProj As clsProjekt = Nothing

                                                If Not firstIteration Then
                                                    ' now start again with hproj with very first startDate 
                                                    hproj = firsthproj
                                                End If

                                                Dim tmpMsg As String = "try out various variants for project .." & hproj.getShapeText
                                                Call logger(ptErrLevel.logInfo, "find best start ", tmpMsg)

                                                Dim endeKriterium1 As Boolean = DateDiff(DateInterval.Day, hproj.startDate, latestEndDate) < minDuration

                                                Do While overutilizationFound And Not endeKriterium1
                                                    ' move project by deltaIndays

                                                    startIterations = startIterations + 1

                                                    If minDuration < maxDuration And Not firstIteration Then


                                                        'Dim endeKriterium2 As Boolean = DateDiff(DateInterval.Day, hproj.startDate.AddDays(hproj.dauerInDays + deltaInDays - 1), latestEndDate) <= 0
                                                        Dim endeKriterium2 As Boolean = hproj.dauerInDays + deltaInDays > maxDuration

                                                        Do While overutilizationFound And Not endeKriterium2

                                                            newEndDate = hproj.endeDate.AddDays(deltaInDays)
                                                            tmpProj = moveProject(hproj, hproj.startDate, newEndDate)

                                                            durationIterations = durationIterations + 1
                                                            sumIterations = sumIterations + 1

                                                            If Not IsNothing(tmpProj) Then

                                                                hproj = tmpProj

                                                                ' protocol ... 
                                                                tmpMsg = "try out: " & hproj.getShapeText & hproj.startDate & " - " & hproj.endeDate
                                                                Call logger(ptErrLevel.logInfo, "status:  ", tmpMsg)

                                                                ' now define showrangeLeft and showrangeRight from hproj 
                                                                showRangeLeft = getColumnOfDate(hproj.startDate)
                                                                showRangeRight = getColumnOfDate(hproj.endeDate)

                                                                If myKennung = PTRpa.visboFindProjectStartPM Then
                                                                    ' aus ShowProjekte rausnehmen, um ReferenzValues zu bestimmen 
                                                                    If ShowProjekte.contains(hproj.name) Then
                                                                        ShowProjekte.Remove(hproj.name, False)
                                                                    End If
                                                                    referenceMSValues = ShowProjekte.getMilestonesFrequency(jobParameters.getMilestoneNames)
                                                                    referencePHValues = ShowProjekte.getPhaseFrequency(jobParameters.getPhaseNames)
                                                                End If

                                                                ' now replace in AlleProjekte, ShowProjekte 
                                                                AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                                                ShowProjekte.AddAnyway(hproj)


                                                                If myKennung = PTRpa.visboFindProjectStart Then
                                                                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                                                                Else
                                                                    overutilizationFound = ShowProjekte.overLoadMSPhasesFound(jobParameters.getMilestoneNames, jobParameters.limitMilestones,
                                                                                                  referenceMSValues,
                                                                                                  jobParameters.getPhaseNames, jobParameters.limitPhases,
                                                                                                  referencePHValues)
                                                                End If



                                                            Else
                                                                ' Error occurred 
                                                                Throw New ArgumentException("tmpProj is Nothing")
                                                            End If

                                                            endeKriterium2 = hproj.dauerInDays + deltaInDays > maxDuration

                                                        Loop

                                                    End If

                                                    If overutilizationFound Then

                                                        newStartDate = hproj.startDate.AddDays(deltaInDays)
                                                        newEndDate = newStartDate.AddDays(minDuration - 1)

                                                        tmpProj = moveProject(hproj, newStartDate, newEndDate)
                                                        ' 

                                                        sumIterations = sumIterations + 1

                                                        If Not IsNothing(tmpProj) Then

                                                            hproj = tmpProj

                                                            tmpMsg = "try out: " & hproj.getShapeText & newStartDate & " - " & newEndDate
                                                            Call logger(ptErrLevel.logInfo, "status:  ", tmpMsg)

                                                            ' now define showrangeLeft and showrangeRight from hproj 
                                                            showRangeLeft = getColumnOfDate(hproj.startDate)
                                                            showRangeRight = getColumnOfDate(hproj.endeDate)

                                                            If myKennung = PTRpa.visboFindProjectStartPM Then
                                                                ' aus ShowProjekte rausnehmen, um ReferenzValues zu bestimmen 
                                                                If ShowProjekte.contains(hproj.name) Then
                                                                    ShowProjekte.Remove(hproj.name, False)
                                                                End If
                                                                referenceMSValues = ShowProjekte.getMilestonesFrequency(jobParameters.getMilestoneNames)
                                                                referencePHValues = ShowProjekte.getPhaseFrequency(jobParameters.getPhaseNames)
                                                            End If


                                                            ' now replace in AlleProjekte, ShowProjekte 
                                                            AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                                            ShowProjekte.AddAnyway(hproj)


                                                            If myKennung = PTRpa.visboFindProjectStart Then
                                                                overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                                                            Else
                                                                overutilizationFound = ShowProjekte.overLoadMSPhasesFound(jobParameters.getMilestoneNames, jobParameters.limitMilestones,
                                                                                                  referenceMSValues,
                                                                                                  jobParameters.getPhaseNames, jobParameters.limitPhases,
                                                                                                  referencePHValues)
                                                            End If

                                                        Else
                                                            ' Error occurred 
                                                            Throw New ArgumentException("tmpProj is Nothing")
                                                        End If
                                                    End If

                                                    endeKriterium1 = DateDiff(DateInterval.Day, hproj.startDate, latestEndDate) < minDuration

                                                Loop

                                            Catch ex As Exception
                                                Dim infomsg As String = "failure: could not create project-variant " & hproj.getShapeText & ex.Message
                                                Call logger(ptErrLevel.logError, "find best start ", infomsg)
                                                overutilizationFound = True
                                            End Try

                                        End If


                                        firstIteration = False
                                    Next

                                    If Not overutilizationFound Then
                                        ' it is already in there ... but now needed to be stored
                                        storeRequired = True
                                    Else
                                        ' take it out again , because there was no solution
                                        storeRequired = False
                                        AlleProjekte.Remove(calcProjektKey(hproj))
                                        ShowProjekte.Remove(hproj.name)
                                    End If

                                Else
                                    ' all ok, just continue
                                End If

                                If storeRequired Then
                                    Dim myMessages As New Collection
                                    If storeSingleProjectToDB(hproj, myMessages) Then
                                        ' now for the sake of sequence in Constellation 
                                        myRowNr = myRowNr + 1

                                        Dim infomsg As String = "success: created " & sumIterations & " variants to avoid bottlenecks " & hproj.getShapeText
                                        Call logger(ptErrLevel.logInfo, "find best start ", infomsg)
                                    Else
                                        ' take it out again , because there was no solution
                                        ShowProjekte.Remove(hproj.name)
                                        Dim infomsg As String = "... failed to store variant to avoid bottlenecks " & hproj.getShapeText
                                        Call logger(ptErrLevel.logError, "find best start ", infomsg)
                                    End If
                                Else
                                    If overutilizationFound Then
                                        ' for showing which projects could not be considered
                                        missingList.Add(hproj)
                                        Dim infomsg As String = "unsuccessful : tried out " & sumIterations & " variants for " & hproj.name
                                        Call logger(ptErrLevel.logWarning, "find best start ", infomsg)
                                    Else
                                        ' now for the sake of sequence in Constellation 
                                        myRowNr = myRowNr + 1

                                        Dim infomsg As String = "success: could be added to portfolio variant as-is " & hproj.getShapeText
                                        Call logger(ptErrLevel.logInfo, "find best start ", infomsg)
                                    End If

                                End If
                            Else
                                Dim infomsg As String = rankingPair.Value.projectName & " does not exist so far"
                                Call logger(ptErrLevel.logError, "find best start ", infomsg)
                            End If
                        End If



                    Next

                    ' now consider to sort it again 
                    If myKennung = PTRpa.visboFindProjectStartPM Then
                        ' do the sorting according the very first phase , or if there is no phase according the first milestones 
                        ' start , then duration 
                        Dim isMilestone As Boolean = False
                        Dim sortItem As String = ""
                        If jobParameters.getPhaseNames.Count > 0 Then
                            sortItem = jobParameters.getPhaseNames.Item(0)
                        ElseIf jobParameters.getMilestoneNames.Count > 0 Then
                            sortItem = jobParameters.getMilestoneNames.Item(0)
                            isMilestone = True
                        End If

                        Dim myNewSortlist As New SortedList(Of Double, String)
                        Try

                            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                                Dim sortKrit As Double = 100000000.0

                                If Not isMilestone Then
                                    ' sort Kriterium is phase start and duration 
                                    Dim elemName As String = ""
                                    Dim breadCrumb As String = ""
                                    Dim lfdNr As Integer = 1
                                    Dim type As Integer = -1
                                    Dim pvName As String = ""
                                    Call splitHryFullnameTo2(sortItem, elemName, breadCrumb, type, pvName)
                                    Dim cPhase As clsPhase = kvp.Value.getPhase(elemName, lfdNr:=1)
                                    sortKrit = DateDiff(DateInterval.Day, StartofCalendar, cPhase.getStartDate) + 1 / cPhase.dauerInDays

                                    Dim kritDelta As Double = 0.00001
                                    Do While myNewSortlist.ContainsKey(sortKrit)
                                        sortKrit = sortKrit + kritDelta
                                    Loop

                                Else
                                    ' sort Kriterium is phase start and duration 
                                    Dim elemName As String = ""
                                    Dim breadCrumb As String = ""
                                    Dim lfdNr As Integer = 1
                                    Dim type As Integer = -1
                                    Dim pvName As String = ""
                                    Call splitHryFullnameTo2(sortItem, elemName, breadCrumb, type, pvName)
                                    Dim cMilestone As clsMeilenstein = kvp.Value.getMilestone(elemName, lfdNr:=1)
                                    sortKrit = DateDiff(DateInterval.Day, StartofCalendar, cMilestone.getDate)

                                    Dim kritDelta As Double = 0.00001
                                    Do While myNewSortlist.ContainsKey(sortKrit)
                                        sortKrit = sortKrit + kritDelta
                                    Loop
                                End If
                                ' jetzt enthält die Sortierte Liste den Eintrag nicht mehr 
                                myNewSortlist.Add(sortKrit, kvp.Value.name)

                            Next

                            ' now re-define the sorting in the constellation
                            Dim tmpSortListe As New SortedList(Of String, String)
                            Dim zeile As Integer = 2
                            Dim formatStr As String = "00000000"
                            For Each kvp2 As KeyValuePair(Of Double, String) In myNewSortlist
                                tmpSortListe.Add(zeile.ToString(formatStr), kvp2.Value)
                                zeile = zeile + 1
                            Next

                            ' jetzt das als die Sort-Liste der Konstellation übernehmen 
                            currentSessionConstellation.sortListe(ptSortCriteria.customTF) = tmpSortListe
                        Catch ex As Exception

                        End Try
                    End If

                    Dim pfVariantName As String = jobParameters.portfolioVariantName & " - " & projectVariantName
                    toStoreConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                                cName:=myActivePortfolio, vName:=pfVariantName)

                    outputCollection.Clear()
                    Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                    If outputCollection.Count > 0 Then
                        Call logger(ptErrLevel.logInfo, "Project List Import, Store Portfolio-Variant: ", outputCollection)
                    End If

                End If

                ' now Store Constellation 
                If missingList.Count > 0 Then
                    Dim portfolioName As String = jobParameters.portfolioName
                    Dim pfVariantName As String = jobParameters.portfolioVariantName & " - " & jobParameters.projectVariantName
                    Dim ok As Boolean = storeConstellationFromProjectList(missingList, portfolioName, pfVariantName & " missing")
                End If
            End If




        Catch ex As Exception
            result = False
        End Try

        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        processProjectListWithActivePortfolio = result

    End Function


    ''' <summary>
    ''' all projects need to be read alrady in ImportProjekte
    ''' </summary>
    ''' <returns></returns>
    Public Function defineFeasiblePortfolio() As Boolean

        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        Dim overloadAllowedinMonths As Double = 1.05
        Dim overloadAllowedTotal As Double = 1.0


        Try
            ' now get the aggregation Roles
            Dim aggregationRoles As SortedList(Of Integer, String) = RoleDefinitions.getAggregationRoles()
            Dim aggregationList As New List(Of String)

            Dim teamID As Integer = -1

            Dim exceptionList As Collection = getConsiderationList(excludedNames:=True)

            ' checkout aggregation Roles
            For Each ar As KeyValuePair(Of Integer, String) In aggregationRoles
                Dim tmpStrID As String = RoleDefinitions.bestimmeRoleNameID(ar.Key, teamID)
                If Not aggregationList.Contains(tmpStrID) And Not exceptionList.Contains(tmpStrID) Then
                    aggregationList.Add(tmpStrID)
                End If
            Next

            ' Get the Ranking out of Excel List , it is just the ordering of the rows 
            ' value holds the AllProjekte.Key, i.e name#variantName
            Dim rankingList As SortedList(Of Integer, String) = getRanking()

            ' takes all the projects which could not be considered first time ... 
            Dim rankingList2 As New SortedList(Of Integer, String)
            Dim missingList As New clsProjekteAlle

            Dim abbruchDate As Date = New Date(2023, 6, 30)

            Dim tmpValues As Double() = getOverloadParams()
            Dim tmpNames As String() = getPortfolioNames()


            Dim portfolioName As String = tmpNames(0)
            Dim variantName As String = tmpNames(1)


            Try
                abbruchDate = CDate(tmpNames(2))
            Catch ex As Exception
                abbruchDate = New Date(2023, 6, 30)
            End Try


            overloadAllowedinMonths = tmpValues(0)
            overloadAllowedTotal = tmpValues(1)


            AlleProjekte.Clear()
            ' now make sure all projects are in AlleProjekte
            For Each ppair As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste

                If Not AlleProjekte.Containskey(ppair.Key) Then
                    AlleProjekte.Add(ppair.Value)
                End If

            Next

            ' then empty ShowProjekte again 
            ShowProjekte.Clear()

            ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
            ' It then represents the sequence: Row1 is the most important project , Row2 the scond, and so forth
            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                ' check whether or not there is already such a project-variant in the DB

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Remove(hproj.name)
                End If

                ShowProjekte.Add(hproj)

                Dim skillIDs As Collection = hproj.getSkillNameIds
                Dim skillList As New List(Of String)

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) And Not exceptionList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next


                ' now consider the tiemFrame this project occupies ..
                showRangeLeft = getColumnOfDate(hproj.startDate)
                showRangeRight = getColumnOfDate(hproj.endeDate)


                Dim overutilizationFound As Boolean = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)

                ' before checking for utilization , do the auto-distribution, that is according to freecapacity try to distribute it in the timeframe that all fits
                ' example: with a normalized distribution it may come to a bootleneck: 15 PT in JAn, 15 PT in Feb. But there is 12 PT in Jan, 18 PT in Feb free capacity , so this works out. 
                Dim myMessages As New Collection
                Dim storeRequired As Boolean = False

                ' tk 24.4. Auto-Distribute : take it out , because it changes too much in the monthly needs 
                'If overutilizationFound Then
                '    Dim infomsg As String = "try out variant with optimized distribution according to resource needs:  " & hproj.getShapeText
                '    Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                '    hproj = hproj.createVariant(variantName, "consider requested sum - distribute according free capacity ")
                '    Dim errorMsg As String = ""

                '    ' put it into AlleProjekte
                '    AlleProjekte.Add(hproj)
                '    ShowProjekte.AddAnyway(hproj)

                '    Call ShowProjekte.autoDistribute(hproj.name, "", errorMsg)


                '    ' now calculate again ... 
                '    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)

                '    ' to trigger saving of the new variant , if there was first a bottleneck, then no more ...
                '    storeRequired = Not overutilizationFound

                'End If


                If overutilizationFound Then

                    ' take it out again, because there was no solution
                    ShowProjekte.Remove(hproj.name)
                    Dim infomsg As String = "with default start-Date & end-Date not considered because of bottlenecks, will be tried out later ... " & hproj.name
                    Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                    'Console.WriteLine(infomsg)

                    rankingList2.Add(rankingPair.Key, rankingPair.Value)

                Else
                    ' all ok, just continue
                    Dim infomsg As String = " ... considered " & hproj.getShapeText
                    Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                    'Console.WriteLine(infomsg)

                    ' now if there was created the variant
                    If storeRequired Then

                        Dim tmpMessages As New Collection
                        If storeSingleProjectToDB(hproj, tmpMessages) Then
                            Dim mymsg As String = "tried out new value distribution:  worked out to find solution for  " & hproj.getShapeText
                            Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", mymsg)

                            'Console.WriteLine(mymsg)

                        End If
                    Else
                        Dim mymsg As String = "could be considered unchanged: " & hproj.getShapeText
                        Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", mymsg)

                        'Console.WriteLine(mymsg)
                    End If
                End If



            Next

            ' now the second wave is going to come ... 


            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList2

                Dim anzLoops As Integer = 1

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                End If

                Dim skillIDs As Collection = hproj.getSkillNameIds
                Dim skillList As New List(Of String)

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) And Not exceptionList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next

                ' now consider the timeFrame this project occupies ..

                Dim key As String = calcProjektKey(hproj)
                ' create variant if not already done
                If hproj.variantName <> variantName Then
                    hproj = hproj.createVariant(variantName, "variant was created and moved to avoid resource bottlenecks")
                    ' bring that into AlleProjekte
                    key = calcProjektKey(hproj)
                    If AlleProjekte.Containskey(key) Then
                        AlleProjekte.Remove(key)
                    End If
                    AlleProjekte.Add(hproj)
                End If

                Dim deltaInDays As Integer = 3

                Dim iterations As Integer = 0

                Dim newStartDate As Date = hproj.startDate.AddDays(deltaInDays)
                Dim newEndDate As Date = hproj.endeDate.AddDays(deltaInDays)

                Dim overutilizationFound As Boolean = True

                Do While overutilizationFound And newEndDate <= abbruchDate
                    ' move project by deltaIndays

                    Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                    If Not IsNothing(tmpProj) Then

                        hproj = tmpProj

                        showRangeLeft = getColumnOfDate(tmpProj.startDate)
                        showRangeRight = getColumnOfDate(tmpProj.endeDate)


                        '' now replace in ShowProjekte 
                        'AlleProjekte.Remove(key)
                        'ShowProjekte.Remove(tmpProj.name)
                        ' add the new, altered version 
                        AlleProjekte.Add(hproj)
                        ShowProjekte.AddAnyway(hproj)

                        overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)

                        ' tk 24.4.22 takte it out again , autoDistribute changes too much
                        If overutilizationFound Then
                            Dim fmsg As String = ""
                            Call ShowProjekte.autoDistribute(hproj.name, "", fmsg)
                            overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)
                        End If

                        If overutilizationFound Then
                            newStartDate = newStartDate.AddDays(deltaInDays)
                            newEndDate = newEndDate.AddDays(deltaInDays)

                            anzLoops = anzLoops + 1
                        End If

                    Else
                        ' Error occurred 
                        Exit Do
                    End If

                Loop

                If Not overutilizationFound Then
                    ' it is already in there ... but now needed to be stored
                    Dim myMessages As New Collection
                    If storeSingleProjectToDB(hproj, myMessages) Then
                        Dim infomsg As String = "tried out " & anzLoops & " different start/ends to avoid bottlenecks, found solution for  " & hproj.getShapeText
                        Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                        'Console.WriteLine(infomsg)

                    Else
                        ' take it out again , because there was no solution
                        ShowProjekte.Remove(hproj.name)

                        Dim mlKEy As String = calcProjektKey(hproj.name, "")
                        hproj = ImportProjekte.getProject(mlKEy)
                        missingList.Add(hproj)

                        Dim infomsg As String = "... failure: could not store " & hproj.getShapeText
                        Call logger(ptErrLevel.logError, "define feasible portfolio: ", infomsg)

                        'Console.WriteLine(infomsg)
                    End If


                Else
                    ' take it out again , because there was no solution
                    AlleProjekte.Remove(key)
                    ShowProjekte.Remove(hproj.name)

                    ' take it out again , because there was no solution
                    Dim infomsg As String = "... could finally not be considered  " & hproj.name

                    Call logger(ptErrLevel.logWarning, "define feasible portfolio: ", infomsg)

                    'Console.WriteLine(infomsg)

                    Dim mlKEy As String = calcProjektKey(hproj.name, "")
                    hproj = ImportProjekte.getProject(mlKEy)
                    missingList.Add(hproj)

                End If

            Next

            ' ----------------


            ' now create the portfolio Variant from ShowProjekte 
            ' now create the Portfolio from ShowProjekte content 

            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=variantName)
            Dim errMsg As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logInfo, "feasible Portfolio definition, ", outputCollection)
            End If

            ' now Store Constellation 
            If missingList.Count > 0 Then
                Dim ok As Boolean = storeConstellationFromProjectList(missingList, portfolioName, variantName & " missing")
            End If

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "feasible Portfolio definition, unexpected failure:", ex.Message)
            result = False
        End Try

        defineFeasiblePortfolio = result

    End Function


    ''' <summary>
    ''' performs creation and optimization when no activePortfolio is defined or does exist
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <param name="overloadAllowedInMonths"></param>
    ''' <param name="overloadAllowedTotal"></param>
    ''' <returns></returns>
    Public Function processProjectListWithoutActivePortfolio(ByVal aggregationList As List(Of String),
                                                              ByVal skillList As List(Of String),
                                                              ByVal portfolioName As String,
                                                              ByVal overloadAllowedInMonths As Double,
                                                              ByVal overloadAllowedTotal As Double) As Boolean
        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        ' cache löschen
        result = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Try
            ' Get the Ranking out of Excel List , it is just the ordering of the rows 
            ' value holds the AllProjekte.Key, i.e name#variantName
            Dim rankingList As SortedList(Of Integer, String) = getRanking()


            AlleProjekte.Clear()
            ' now make sure all projects are in AlleProjekte
            For Each ppair As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste
                If Not AlleProjekte.Containskey(ppair.Key) Then
                    AlleProjekte.Add(ppair.Value)
                End If
            Next


            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                End If
            Next


            ' currentSessionConstellation is build by alle the Showprojekte.add and AlleProjekte.add Commands ...
            ' create form that a portfolio, only containing the show-Elements 
            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:="")

            ' now store the Portfolio , with name portfolioName
            Dim errMsg As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            ' define the range, necessary to check whether or not there are bottlenecks 
            showRangeLeft = ShowProjekte.getMinMonthColumn
            showRangeRight = ShowProjekte.getMaxMonthColumn


            ' then empty ShowProjekte again 
            ShowProjekte.Clear()


            ' 1. now start with the (next-)highest ranked project, 
            ' 2. If there are no bottlenecks, keep it in ShowProjekte; 
            '    if there are bottlenecks create a variant with name [arb], then move variant by 7 days until there is no bottleneck any more or until project has been moved by approx 6 months
            '    if bottleneck cannot be solved, take project out of potential portfolio 
            ' 3. Go to 1.


            ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
            ' It then represents the sequence: Row1 is the most important project 
            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                End If

                Dim overutilizationFound As Boolean = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)


                If overutilizationFound Then
                    Dim key As String = calcProjektKey(hproj)
                    ' create variant if not already done
                    If hproj.variantName <> "arb" Then
                        hproj = hproj.createVariant("arb", "variant was created and moved to avoid resource bottlenecks")
                        ' bring that into AlleProjekte
                        key = calcProjektKey(hproj)
                        If AlleProjekte.Containskey(key) Then
                            AlleProjekte.Remove(key)
                        End If
                        AlleProjekte.Add(hproj)
                    End If

                    Dim deltaInDays As Integer = 7
                    Dim maxIterations As Integer = CInt(182 / deltaInDays)
                    Dim iterations As Integer = 0

                    Do While overutilizationFound And iterations <= maxIterations
                        ' move project by deltaIndays

                        Dim newStartDate As Date = hproj.startDate.AddDays(deltaInDays)
                        Dim newEndDate As Date = hproj.endeDate.AddDays(deltaInDays)

                        Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                        If Not IsNothing(tmpProj) Then

                            hproj = tmpProj

                            ' now replace in ShowProjekte 
                            AlleProjekte.Remove(key)
                            ShowProjekte.Remove(tmpProj.name)
                            ' add the new, altered version 
                            AlleProjekte.Add(tmpProj)
                            ShowProjekte.Add(tmpProj)

                            overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)

                            If overutilizationFound Then
                                iterations = iterations + 1
                            End If

                        Else
                            ' Error occurred 
                            Exit Do
                        End If

                    Loop

                    If Not overutilizationFound Then
                        ' it is already in there ... but now needed to be stored
                        Dim myMessages As New Collection
                        If storeSingleProjectToDB(hproj, myMessages) Then
                            Dim infomsg As String = "created variant to avoid bottlenecks " & hproj.getShapeText
                            Call logger(ptErrLevel.logInfo, infomsg, myMessages)
                            'Console.WriteLine(infomsg)
                        Else
                            ' take it out again , because there was no solution
                            ShowProjekte.Remove(hproj.name)
                            Dim infomsg As String = "... failed to create variant to avoid bottlenecks " & hproj.getShapeText
                            Call logger(ptErrLevel.logError, infomsg, myMessages)
                            'Console.WriteLine(infomsg)
                        End If


                    Else
                        ' take it out again , because there was no solution
                        AlleProjekte.Remove(key)
                        ShowProjekte.Remove(hproj.name)
                    End If

                Else
                    ' all ok, just continue
                End If

            Next

            ' now create the portfolio Variant arb from ShowProjekte 
            ' now create the Portfolio from ShowProjekte content 

            toStoreConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:="arb")


            outputCollection.Clear()
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logError, "Project List Import, Store Portfolio-Variant arb failed:", outputCollection)
            End If

        Catch ex As Exception
            result = False
        End Try

        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        processProjectListWithoutActivePortfolio = result


    End Function

    ''' <summary>
    ''' creates a quality check file for all bhtc projects
    ''' </summary>
    ''' <param name="otherPortfolios"></param>
    ''' <param name="myPortfolioName"></param>
    ''' <param name="myPortfolioVName"></param>
    ''' <param name="myTemplate"></param>
    Public Sub writeBHTCQualityCheck(ByVal otherPortfolios As String(), ByVal myPortfolioName As String,
                                     Optional ByVal myPortfolioVName As String = "",
                                     Optional ByVal myTemplate As String = "TMS")

        Dim portfolio As clsConstellation = Nothing
        Dim Err As clsErrorCodeMsg = Nothing
        Dim allOK As Boolean = True
        Dim tmpID As String = ""
        Dim expFName As String = ""



        Dim reportWB As xlns.Workbook = Nothing

        ' now get Portfolio from VISBO cloud 
        Dim heute As Date = Date.Now
        Dim pvName As String = calcPortfolioKey(myPortfolioName, myPortfolioVName)

        Dim compareTemplate As clsProjektvorlage = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectTemplatefromDB(myTemplate, tmpID, heute, Err)

        Dim myConstellations As clsConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(heute, Err)

        If myConstellations.Contains(pvName) And Not IsNothing(compareTemplate) Then

            Dim myConstellation As clsConstellation = myConstellations.getConstellation(myPortfolioName, myPortfolioVName)
            ' if successful: create / open Excel Export File 

            expFName = exportOrdnerNames(PTImpExp.rpa) & "\" & "Quality Check " & myConstellation.constellationName & ".xlsx"


            ' hier muss jetzt das entsprechende File aufgemacht werden ...
            ' das File 
            Try

                reportWB = appInstance.Workbooks.Add()
                CType(reportWB.Worksheets.Item(1), xlns.Worksheet).Name = "VISBO"
                reportWB.SaveAs(Filename:=expFName, ConflictResolution:=xlns.XlSaveConflictResolution.xlLocalSessionChanges)

            Catch ex As Exception
                Call logger(ptErrLevel.logError, "Creating Excel File Output File", " failed ..")
                appInstance.EnableEvents = True
                allOK = False
            End Try

            If allOK Then
                Dim ws As xlns.Worksheet = CType(reportWB.Worksheets("VISBO"), xlns.Worksheet)

                ' now write Headerline 
                Dim zeile As Integer = 1
                Dim spalte As Integer = 1
                ws.Cells(zeile, 1).value = "Project-Name"
                ws.Cells(zeile, 2).value = "VISBO ID"
                ws.Cells(zeile, 3).value = "State"
                ws.Cells(zeile, 4).value = "Start Date"
                ws.Cells(zeile, 5).value = "End Date"
                ws.Cells(zeile, 6).value = "Inner Start "
                ws.Cells(zeile, 7).value = "Inner End "
                ws.Cells(zeile, 8).value = "Inner End "
                ws.Cells(zeile, 9).value = "Current?"
                ws.Cells(zeile, 10).value = "%Done Quality"
                ws.Cells(zeile, 11).value = "last Publish"
                ws.Cells(zeile, 12).value = "Comparability Index"
                ws.Cells(zeile, 13).value = "is Part of Portfolio"
                ws.Cells(zeile, 14).value = "is Part of other Portfolios"


                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    zeile = zeile + 1
                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    ' read it , but do not put into AlleProjekte 
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)

                    Dim innerStartEndDate() As Date = hproj.getInnerStartEndDate

                    Dim myState As String = hproj.vpStatus


                    ' now writing 
                    ws.Cells(zeile, 1).value = hproj.name
                    ws.Cells(zeile, 2).value = hproj.vpID
                    ws.Cells(zeile, 3).value = hproj.vpStatus
                    ws.Cells(zeile, 4).value = hproj.startDate.ToShortDateString
                    ws.Cells(zeile, 5).value = hproj.endeDate.ToShortDateString
                    ws.Cells(zeile, 6).value = innerStartEndDate(0)
                    ws.Cells(zeile, 7).value = innerStartEndDate(1)

                    ' now check whether it seems to be an active project 
                    Dim isActive As Boolean = DateDiff(DateInterval.Day, innerStartEndDate(0), heute) >= 0 And
                                                DateDiff(DateInterval.Day, innerStartEndDate(1), heute) <= 0

                    If isActive Then
                        ws.Cells(zeile, 8).value = "Yes"
                    Else
                        ws.Cells(zeile, 8).value = "No"
                    End If

                    ' now check whether it complies to TMS structure
                    Dim hasTMS As Boolean = hproj.hasStructureOf(compareTemplate)

                    If hasTMS Then
                        ws.Cells(zeile, 9).value = "Yes"
                    Else
                        ws.Cells(zeile, 9).value = "No"
                    End If

                    ' check the %-Done Quality of Past Elements : Past meaning elements before hproj.timestamp
                    Dim doneQualityKPI As Double = hproj.getdoneQualityKPI()
                    ws.Cells(zeile, 10).value = doneQualityKPI

                    ' Check the last publish - again an indicator of how reliable data is ... 
                    ws.Cells(zeile, 11).value = hproj.timeStamp

                    'check the "omparability Index: keep the current and compare with former versions. How many elements of former versions are in current version ? 
                    Dim resultString As String = ""
                    Dim lookForTimeStamp As Date = hproj.timeStamp.AddMonths(-1)
                    Dim compareVersion As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, hproj.vpID, lookForTimeStamp, Err)

                    Do While Not IsNothing(compareVersion)

                        If resultString = "" Then
                            resultString = hproj.getCompareKPI(compareVersion).ToString("#.##")
                        Else
                            resultString = resultString & " / " & hproj.getCompareKPI(compareVersion).ToString("#.##")
                        End If

                        lookForTimeStamp = compareVersion.timeStamp.AddMonths(-1)
                        compareVersion = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, hproj.vpID, lookForTimeStamp, Err)
                    Loop

                    ws.Cells(zeile, 12).value = resultString

                    ' now check in which portfolio that project is in 
                    'ws.Cells(zeile, 12).value = "is Part of Portfolio"
                    'ws.Cells(zeile, 13).value = "is Part of other Portfolios"
                    Dim containedIn As String = ""
                    Dim containedAlsoIn As String = ""
                    Dim first As Boolean = True

                    For Each pfKVP As KeyValuePair(Of String, clsConstellation) In myConstellations.Liste
                        ' kvp.key now contains the pvName of currently considered project
                        If pfKVP.Value.contains(kvp.Key, False) Then
                            If first Then
                                containedIn = pfKVP.Value.constellationName
                                first = False
                            Else
                                If containedAlsoIn = "" Then
                                    containedAlsoIn = pfKVP.Value.constellationName
                                Else
                                    containedAlsoIn = containedAlsoIn & "; " & pfKVP.Value.constellationName
                                End If
                            End If
                        End If
                    Next

                    ws.Cells(zeile, 13).value = containedIn
                    ws.Cells(zeile, 14).value = containedAlsoIn

                Next

            End If

        Else
            Dim msgTxt As String = "Load Portfolio " & myPortfolioName & " failed .."
            Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
            allOK = False
        End If


        Try
            ' jetzt die Autofilter aktivieren ... 
            If Not CType(reportWB.Worksheets("VISBO"), xlns.Worksheet).AutoFilterMode = True Then
                CType(reportWB.Worksheets("VISBO"), xlns.Worksheet).Cells(1, 1).AutoFilter()
            End If

            reportWB.Close(SaveChanges:=True)
            Call logger(ptErrLevel.logInfo, "Quality Check Successful, stored in ", expFName)
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Store Excel File ", " failed ..")
        End Try

        appInstance.EnableEvents = True

    End Sub



End Module
