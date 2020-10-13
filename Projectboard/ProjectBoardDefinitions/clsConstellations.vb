Public Class clsConstellations
    Private _allConstellations As SortedList(Of String, clsConstellation)
    ' der bool'sche Wert hat aktuell keine Bedeutung; später evtl benutzen um zu bestimmen, das Portfolio Budget aus den Einzelbudgets zu berechnen ist
    Private _listOfLoadedSessionPortfolios As SortedList(Of String, Boolean)

    Public ReadOnly Property Count As Integer

        Get
            Count = _allConstellations.Count
        End Get

    End Property

    ''' <summary>
    ''' returns number of loaded portfolios
    ''' 0 , if no portfolio has been loaded so far
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CountLoadedPortfolios As Integer
        Get
            CountLoadedPortfolios = _listOfLoadedSessionPortfolios.Count
        End Get
    End Property

    ''' <summary>
    ''' fügt der Liste an loadedSessionPortfolios ein Portfolio hinzu 
    ''' wenn zum Portfolio noch kein Summary Projekt existiert, wird es erstellt 
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <returns></returns>
    Public Function addToLoadedSessionPortfolios(ByVal portfolioName As String, Optional ByVal vName As String = "") As Boolean
        Dim tmpResult As Boolean = True
        Dim pvName As String = calcPortfolioKey(portfolioName, vName)
        If Me.Contains(pvName) Then
            If Not _listOfLoadedSessionPortfolios.ContainsKey(pvName) Then
                _listOfLoadedSessionPortfolios.Add(pvName, True)
            End If
        Else
            tmpResult = False
        End If
        ' was ist mit dem entsprechenden Summary Projekt ... 
        Dim tmpVariantName As String = ""
        If awinSettings.loadPFV Then
            tmpVariantName = ptVariantFixNames.pfv.ToString
        End If
        Dim skey As String = calcProjektKey(portfolioName, tmpVariantName)

        Dim hproj As clsProjekt = AlleProjekte.getProject(key:=skey)
        If Not IsNothing(hproj) Then
            If AlleProjektSummaries.Containskey(skey) Then
                AlleProjektSummaries.Remove(skey, False)
            End If
            AlleProjektSummaries.Add(hproj, updateCurrentConstellation:=False, checkOnConflicts:=False)
        End If

        addToLoadedSessionPortfolios = tmpResult
    End Function

    ''' <summary>
    ''' setzt die Liste der geladenen Session Portfolios zurück 
    ''' </summary>
    Public Sub clearLoadedPortfolios()
        _listOfLoadedSessionPortfolios.Clear()
        AlleProjektSummaries.Clear(False)
    End Sub

    ''' <summary>
    ''' setzt die Liste aller Portfolios der Session zurück 
    ''' </summary>
    Public Sub Clear()
        _listOfLoadedSessionPortfolios.Clear()
        _allConstellations.Clear()
        AlleProjektSummaries.Clear(False)
    End Sub

    ''' <summary>
    ''' gibt das Gesamt Budget des Zeitraums im Gesamt-Portfolio zurück 
    ''' noch TODO
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getBudgetOfLoadedPortfolios() As Double
        Get
            Dim tmpResult As Double = 0.0
            For Each pfKvP As KeyValuePair(Of String, Boolean) In _listOfLoadedSessionPortfolios

                Dim variantName As String = ""
                If awinSettings.loadPFV Then
                    variantName = ptVariantFixNames.pfv.ToString
                End If
                Dim key As String = calcProjektKey(pfKvP.Key, variantName)
                Dim hproj As clsProjekt = AlleProjekte.getProject(key)

                ' wenn ein Portfolio nicht über Platzhalter geladen wird, dann werden die Summary Projekte in AlleProjektSummaries platziert ..
                If IsNothing(hproj) Then
                    hproj = AlleProjektSummaries.getProject(key)
                End If

                Dim teilbudget As Double, pk As Double, sk As Double, rk As Double, ergebnis As Double

                If Not IsNothing(hproj) Then
                    Call hproj.calculateRoundedKPI(teilbudget, pk, sk, rk, ergebnis)
                    tmpResult = tmpResult + teilbudget
                End If

            Next
            getBudgetOfLoadedPortfolios = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' liefert die Namen der Szenarien zurück, getrennt durch ;, die pName und vName referenzieren 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="vname"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSzenarioNamesWith(ByVal pname As String, ByVal vname As String,
                                                  Optional ByVal inclExplanation As Boolean = True) As String
        Get

            Dim initMSg As String = "referenced by Portfolio(s: " & vbLf
            Dim tmpResult As String = ""

            If Not inclExplanation Then
                initMSg = ""
            End If

            For Each kvp As KeyValuePair(Of String, clsConstellation) In _allConstellations

                If kvp.Key = "Last" Or kvp.Key = "Session" Then
                    ' nichts tun, sind systembedingt 
                Else
                    If vname = "$ALL" Then
                        If kvp.Value.containsProject(pname) Then
                            If tmpResult = "" Then
                                tmpResult = initMSg & kvp.Key
                            Else
                                tmpResult = tmpResult & "; " & kvp.Key
                            End If
                        End If
                    Else
                        Dim tmpKey As String = calcProjektKey(pname, vname)
                        If kvp.Value.contains(tmpKey, False) Then
                            If tmpResult = "" Then
                                tmpResult = initMSg & kvp.Key
                            Else
                                tmpResult = tmpResult & "; " & kvp.Key
                            End If
                        End If
                    End If

                End If
            Next

            getSzenarioNamesWith = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob eine der Constellations in der Collection einen Konflikt mit der angegebenen Constellation hat
    ''' Konflikt heisst: gleiches Projekt referenziert, egal welche Variante 
    ''' </summary>
    ''' <param name="otherConstellation"></param>
    ''' <returns></returns>
    Public Function hasAnyConflictsWith(ByVal otherConstellation As clsConstellation) As Boolean
        Dim i As Integer = 0
        Dim hasConflict As Boolean = False

        Do While i < Count And Not hasConflict
            Dim aktConst As clsConstellation = _allConstellations.ElementAt(i).Value
            If aktConst.hasAnyConflictsWith(otherConstellation) Then
                hasConflict = True
            Else
                i = i + 1
            End If
        Loop

        hasAnyConflictsWith = hasConflict
    End Function
    Public ReadOnly Property Liste As SortedList(Of String, clsConstellation)

        Get
            Liste = _allConstellations
        End Get

    End Property

    Public ReadOnly Property getConstellation(name As String, Optional ByVal vname As String = "") As clsConstellation

        Get
            If name <> "" Then
                Dim pvname As String = calcPortfolioKey(name, vname)
                If _allConstellations.ContainsKey(pvname) Then
                    getConstellation = _allConstellations.Item(pvname)
                Else
                    getConstellation = Nothing
                End If
            Else
                getConstellation = Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Contains(name As String) As Boolean
        Get
            Contains = _allConstellations.ContainsKey(name)
        End Get
    End Property

    ''' <summary>
    ''' trägt die angegebene Constellation in die Liste der ProjektConstellations ein. 
    ''' </summary>
    ''' <param name="item">Constellation</param>
    Sub Add(ByVal item As clsConstellation)
        Dim trennzeichen = "#"
        Try
            _allConstellations.Add(item.constellationName & trennzeichen & item.variantName, item)
        Catch ex As Exception
            Dim errmsg = "Program-/Portfolio Name existiert bereits: " & item.constellationName & "[" & item.variantName & "]"
            Throw New ArgumentException(errmsg)
        End Try


    End Sub


    ''' <summary>
    ''' aktualisiert in jeder Constellation die Variante mit Namen oldvName mit dem Namen newvName
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="oldvName"></param>
    ''' <param name="newvName"></param>
    ''' <remarks></remarks>
    Public Sub updateVariantName(ByVal pName As String, ByVal oldvName As String, ByVal newvName As String)

        For Each kvp As KeyValuePair(Of String, clsConstellation) In _allConstellations
            Call kvp.Value.updateVariantName(pName, oldvName, newvName)
        Next

    End Sub


    ''' <summary>
    ''' ersetzt oder fügt eine neue Konstellation mit dem Namen ein 
    ''' das wird die neue Konstellation , das heisst es muss auch ein neues Summary Projekt für diese Konstellation gemacht werden  
    ''' </summary>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Public Sub update(item As clsConstellation)
        Dim trennzeichen As String = "#"
        Me.clearLoadedPortfolios()
        Dim pvName As String = calcPortfolioKey(item)
        If Me._allConstellations.ContainsKey(pvName) Then
            Me._allConstellations.Remove(pvName)
        End If

        Me._allConstellations.Add(pvName, item)

        ' jetzt das in loadedSessionPortfolios reinbringen 
        Me.addToLoadedSessionPortfolios(item.constellationName, item.variantName)

    End Sub

    Sub Remove(ByVal key As String)

        Try
            _allConstellations.Remove(key)
        Catch ex As Exception
            Throw New ArgumentException("Konstellation" & " key " & "konnte nicht gelöscht werden ")
        End Try

    End Sub

    Sub New()

        _allConstellations = New SortedList(Of String, clsConstellation)
        _listOfLoadedSessionPortfolios = New SortedList(Of String, Boolean)

    End Sub

End Class
