Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports DBAccLayer
Module testMethoden


    ''' <summary>
    ''' schreibt alle Projekte in die Datenbank, liest sie und vergleicht sie auf Identität
    ''' Voraussetzung: showrangeleft , right mus sgesetzt sein 
    ''' </summary>
    ''' <param name="testProjekte"></param>
    ''' <returns></returns>
    Public Function ReadWriteRoundTrip(ByVal testProjekte As clsProjekteAlle) As Collection
        Dim outputCollection As New Collection

        ReadWriteRoundTrip = outputCollection
    End Function

    Public Function testRoleNames() As Collection
        Dim hproj As clsProjekt
        Dim outputCollection As New Collection
        Dim errmsg As String = ""

        Dim atleastOne As Boolean = False


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        ' Testreihe 1: sind die angegebenen Rollen identisch ? 
        Dim duration1 As Integer = 0
        Dim duration2 As Integer = 0


        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

            hproj = kvp.Value

            Dim usedRollen1 As Collection = hproj.getRoleNames
            Dim usedRollen2 As Collection = hproj.rcLists.getRoleNames

            ' Test auf Identität der beiden usedRollen1,2

            If usedRollen1.Count <> usedRollen2.Count Then
                atleastOne = True
            Else
                For ix As Integer = 1 To usedRollen1.Count
                    If Not usedRollen2.Contains(CStr(usedRollen1.Item(ix))) Then
                        Dim name1 As String = CStr(usedRollen1.Item(ix))
                        Dim name2 As String = CStr(usedRollen2.Item(ix))
                        atleastOne = True
                    End If

                Next
            End If

            Dim usedRollen3 As Collection = hproj.getRoleNameIDs
            Dim usedRollen4 As Collection = hproj.rcLists.getRoleNameIDs

            ' Test auf Identität der beiden usedRollen1,2

            If usedRollen3.Count <> usedRollen4.Count Then
                atleastOne = True
            Else
                For ix As Integer = 1 To usedRollen3.Count
                    If Not usedRollen4.Contains(CStr(usedRollen3.Item(ix))) Then
                        Dim name1 As String = CStr(usedRollen4.Item(ix))
                        Dim name2 As String = CStr(usedRollen3.Item(ix))
                        atleastOne = True
                    End If
                Next
            End If


            Dim usedCost1 As Collection = hproj.getCostNames
            Dim usedCost2 As Collection = hproj.rcLists.getCostNames

            If usedCost1.Count <> usedCost2.Count Then
                atleastOne = True
            Else
                For ix As Integer = 1 To usedCost1.Count
                    If Not usedCost2.Contains(CStr(usedCost1.Item(ix))) Then
                        Dim name1 As String = CStr(usedCost1.Item(ix))
                        Dim name2 As String = CStr(usedCost2.Item(ix))
                        atleastOne = True
                    End If

                Next
            End If


        Next

        If atleastOne Then
            errmsg = "bei Rollen/Kosten nicht alles ok ..."
            outputCollection.Add(errmsg)
        Else
            errmsg = "bei Rollen/Kosten alles ok .."
            outputCollection.Add(errmsg)
        End If


        ' Test-Zyklus 2 
        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' alte Methode .... 

            ' mach es möglichst oft ...

            For iter As Integer = 1 To 1

                'For ix As Integer = 1 To RoleDefinitions.Count
                '    Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(ix)

                '    Dim zeitraumBedarf() As Double = ShowProjekte.getRoleValuesInMonth(role.UID.ToString, True)
                '    Dim zeitraumBedarf2() As Double = ShowProjekte.getRoleValuesInMonth(role.UID.ToString, True)

                '    If arraysAreDifferent(zeitraumBedarf, zeitraumBedarf2) Then
                '        atleastOne = True
                '    End If

                'Next

                'If atleastOne Then
                '    Call MsgBox("Rollen-Summen nicht alles ok ...")
                'Else
                '    Call MsgBox("Rollen-Summen alles ok ..")
                'End If
                atleastOne = False

                For ix As Integer = 1 To CostDefinitions.Count
                    Dim cost As clsKostenartDefinition = CostDefinitions.getCostdef(ix)

                    Dim zeitraumBedarf() As Double = ShowProjekte.getCostValuesInMonth(cost.UID)
                    Dim zeitraumBedarf2() As Double = ShowProjekte.getCostValuesInMonthNew(cost.UID)

                    If arraysAreDifferent(zeitraumBedarf, zeitraumBedarf2) Then
                        atleastOne = True
                    End If
                Next

                If atleastOne Then
                    errmsg = "bei Rollen/Kosten nicht alles ok ..."
                    outputCollection.Add(errmsg)
                Else
                    errmsg = "bei Rollen/Kosten alles ok .."
                    outputCollection.Add(errmsg)
                End If

                If atleastOne Then
                    Call MsgBox("Kosten-Summen nicht alles ok ...")
                Else
                    Call MsgBox("Kosten-Summen alles ok ..")
                End If

            Next

            'Call MsgBox("jetzt ist es: " & Date.Now.ToLongTimeString)

            'For iter As Integer = 1 To 400

            '    For ix As Integer = 1 To RoleDefinitions.Count
            '        Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(ix)

            '        Dim zeitraumBedarf() As Double = ShowProjekte.getRoleValuesInMonthNew(role.UID, True)

            '    Next
            'Next

            Call MsgBox("jetzt ist es: " & Date.Now.ToLongTimeString)

        Else
            Call MsgBox("zuerst Zeitraum definieren ...")
        End If



        enableOnUpdate = True

        testRoleNames = outputCollection

    End Function

    ''' <summary>
    ''' macht aus allen Projekten aus ShowProjekte die entsprechenden Aggregate-Projekte
    ''' sollte dann aufgerufen werden, wenn ein Gruppen-Manager  Manager Batch-Datei importiert hat 
    ''' </summary>
    ''' <returns></returns>
    Public Function TestAggregateMethod() As Collection

        Dim outputCollection As New Collection
        Dim found As Boolean = False
        Dim testRoleIDs() As Integer = Nothing
        Dim testportfolioProjekte As New clsProjekte
        Dim array1() As Double
        Dim array2() As Double
        Dim aggregatedProject As clsProjekt = Nothing
        Dim errMsg As String = ""
        Dim ix As Integer = 0

        If Not showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            Call MsgBox("zuerst Showrange belegen ...")
            TestAggregateMethod = outputCollection
            Exit Function
        End If

        For Each kvp As KeyValuePair(Of String, clsCustomUserRole) In customUserRoles.liste

            If kvp.Value.customUserRole = ptCustomUserRoles.PortfolioManager Then
                found = True
                testRoleIDs = kvp.Value.getAggregationRoleIDs
                Exit For
            End If

        Next

        Dim chckSum1() As Double
        Dim chckSum2() As Double
        ReDim chckSum1(testRoleIDs.Length - 1)
        ReDim chckSum2(testRoleIDs.Length - 1)

        If found Then

            ' komplettes Showprojekte überprüfen ...

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                aggregatedProject = kvp.Value.aggregateForPortfolioMgr(testRoleIDs)

                ix = 0
                For Each roleID As Integer In testRoleIDs
                    array1 = kvp.Value.getRessourcenBedarf(roleID, True, False, False)
                    array2 = aggregatedProject.getRessourcenBedarf(roleID, True, False, False)

                    If arraysAreDifferent(array1, array2) Then
                        errMsg = "Unterschiede Projekt / Aggregated Projekt: )" & kvp.Value.name
                        outputCollection.Add(errMsg)
                    End If

                    chckSum1(ix) = chckSum1(ix) + array1.Sum
                    chckSum2(ix) = chckSum2(ix) + array2.Sum
                    ix = ix + 1
                Next


                testportfolioProjekte.Add(kvp.Value, False)


            Next

            For Each roleID As Integer In testRoleIDs

                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(roleID, -1)
                array1 = ShowProjekte.getRoleValuesInMonth(roleNameID, True, PTcbr.all, Nothing)
                array2 = testportfolioProjekte.getRoleValuesInMonth(roleNameID, True, PTcbr.all, Nothing)

                If arraysAreDifferent(array1, array2) Then
                    errMsg = "Unterschiede Portfolio / Aggregated Portfolio: )" & array1.Sum.ToString & " vs. " & array2.Sum.ToString
                    outputCollection.Add(errMsg)
                End If

                If chckSum1.Sum <> array1.Sum Then
                    errMsg = "Unterschiede Summe Einzelprojekte / Portfolio: )" & chckSum1.Sum.ToString & " vs. " & array1.Sum.ToString
                    outputCollection.Add(errMsg)
                End If

                If chckSum2.Sum <> array2.Sum Then
                    errMsg = "Unterschiede Summe Einzelprojekte / Portfolio: )" & chckSum2.ToString & " vs. " & array2.Sum.ToString
                    outputCollection.Add(errMsg)
                End If


            Next

        End If

        TestAggregateMethod = Nothing


    End Function


End Module
